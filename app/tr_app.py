# -*- coding: utf-8 -*-
# TR_APP_MARKER — ห้ามลบบรรทัดนี้ ตัวเปิดใช้เช็กว่าโหลดไฟล์ถูกตัว
"""
TR Plus Ultra — โปรแกรมหลัก
===========================
V0.6.3 : หน้าค้นหากางดูรายการที่จะค้นหาได้

ไฟล์นี้อยู่บน GitHub ตัวเปิด (.exe) จะโหลดมารันทุกครั้ง
แก้ไฟล์นี้แล้ว push = ทุกคนได้ของใหม่ทันที ไม่ต้อง build .exe ใหม่
"""
import os
import re
import csv
import shutil
import json
import asyncio
import queue
import threading
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog

from playwright.async_api import async_playwright

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.worksheet.datavalidation import DataValidation
    XLSX_OK = True
except Exception:
    XLSX_OK = False


# ============================================================================
#  [1] CONFIG — ทุกอย่างที่ผูกกับหน้าตาเว็บอยู่ตรงนี้ก้อนเดียว
#      เว็บเปลี่ยนเมื่อไหร่ แก้แค่บล็อกนี้
#      (ค่าทั้งหมดยืนยันกับเว็บจริงแล้ว 2026-09-07)
# ============================================================================
BASE = 'https://aztek-tools-v2.thehof.gg'
ITEM_LIST_URL = BASE + '/hof/talesrunner/shop/items'

SEL = {
    # หน้า list
    'search_input': 'input[placeholder*="ค้นหา Item"]',
    'search_button': 'button:has-text("ค้นหา")',
    'filter_option': '#filter-option',      # ItemOption
    'filter_duration': '#filter-duration',  # DurationIndex
    'row': 'table tbody tr',
    'page_input': 'input[type="number"]',
    'next_text': 'Next',
    # หน้ารายละเอียด
    'd_name': '#item-name',
    'd_kind': '#item-game-item-id',
    'd_price': '#item-option',
    'd_duration': '#duration-index',
    'd_qty': '#display-quantity',
    'd_trade_label': 'แลกเปลี่ยน',
}

# ---- หน้าสร้างไอเทม ----
ITEM_CREATE_URL = BASE + '/hof/talesrunner/shop/items/create'
SEL_CREATE = {
    # ช่องกรอก (id ของเว็บ v2)
    'name':     '#item-name',            # ชื่อไอเทม
    'kind':     '#item-game-item-id',    # game_item_id (ItemKind / fdItemNum)
    'desc':     '#item-detail',          # คำอธิบาย
    'price':    '#item-option',          # Price
    'duration': '#duration-index',       # ระยะเวลาไอเทม (วัน)
    'mail':     '#mail-title',           # หัวข้อจดหมาย
    'qty':      '#display-quantity',     # จำนวน (ส่วนแสดงบนเว็บ)
    # ตัวที่ไม่มี id แน่นอน -> หาโดยอ้าง "ข้อความที่อยู่ข้างๆ"
    'type_label':  'ประเภทไอเทม',
    'type_ph':     'เลือกประเภท',
    'web_label':   'เปิดใช้งานการแสดงผลบนเว็บ',
    'trade_label': 'แลกเปลี่ยนได้',
    'submit':      'สร้าง Item',
}
DEFAULT_TYPE = 'GENERAL'
DEFAULT_SUFFIX = '(Gift)'
# ลำดับคอลัมน์ในตาราง: Aztek Item Id | ชื่อ | ประเภท | ItemKind | Actions
COL = {'id': 0, 'name': 1, 'type': 2, 'kind': 3}

TMPL_HEADERS = ['ItemKind', 'ItemName', 'duration', 'trade', 'qty']

# ---- ชื่อคอลัมน์ที่ยอมรับได้ (ชีทแต่ละใบใช้ชื่อไม่เหมือนกัน) ----
# เรียงตามลำดับความสำคัญ ตัวแรกที่เจอจะถูกเลือกเป็นค่าเริ่มต้น
ALIAS = {
    'kind': ['itemkind', 'fditemnum', 'item kind', 'fditemid', 'itemid', 'fditemkind'],
    'name': ['itemname', 'display name', 'displayname', 'ชื่อไอเทม', 'ชื่อ', 'item name'],
    'dur':  ['duration', 'ระยะเวลาของขวัญ', 'ระยะเวลา', 'durationindex', 'วัน'],
    'trade': ['trade', 'แลกเปลี่ยน', 'แลกเปลี่ยนได้', 'tradeable'],
    'qty':  ['qty', 'quantity', 'จำนวน'],
}
# แถวที่ถือว่าเป็น "หัวตาราง" ต้องมีคำใดคำหนึ่งนี้
HEADER_MARKERS = ('itemkind', 'fditemnum', 'fditemkind')
SCAN_ROWS = 400        # ลึกสุดที่จะไล่หาหัวตารางในแต่ละชีท

APP_VERSION = globals().get('TRPU_VERSION') or 'dev'
DATA_DIR = globals().get('TRPU_DATA_DIR') or os.path.join(
    os.environ.get('LOCALAPPDATA') or os.path.expanduser('~'), 'TRPlusUltra')
os.makedirs(DATA_DIR, exist_ok=True)
CHROME_PROFILE = os.path.join(DATA_DIR, 'chrome_profile')
PREF_FILE = os.path.join(DATA_DIR, 'prefs.json')

# ---- บันทึกการใช้งาน ----
# เก็บไว้ในโฟลเดอร์ข้อมูลโปรแกรมเป็นค่าเริ่มต้น แต่ย้ายที่ได้จากในโปรแกรม
LOCAL_LOG_DIR = DATA_DIR
LOG_MAX_EVENTS = 20000          # เกินนี้จะตัดของเก่าทิ้ง

# โฟลเดอร์กลางของทีม — เว้นว่าง = ปิดอยู่ (เก็บเฉพาะในเครื่อง)
# พร้อมเมื่อไหร่ ใส่ path ที่ทุกคนเขียนได้ เช่น
#   CENTRAL_LOG_DIR = r'C:\Users\B_A_N\OneDrive\TR-Logs'
# แล้วโปรแกรมจะเขียน log ซ้ำอีกชุดเป็นไฟล์ของแต่ละเครื่องลงไปให้เอง
CENTRAL_LOG_DIR = ''

TR_USER = os.environ.get('USERNAME') or os.environ.get('USER') or 'unknown'
TR_MACHINE = os.environ.get('COMPUTERNAME') or 'unknown'

C = {
    'bg': '#11161f', 'card': '#151c28', 'line': '#263041', 'input': '#0c1119',
    'fg': '#dfe6f2', 'dim': '#8b98ad', 'accent': '#4f6bed',
    'ok': '#5ed09a', 'warn': '#e3b341', 'err': '#f0736a',
}
FM = ('Segoe UI', 10)
FB = ('Segoe UI', 10, 'bold')


# ============================================================================
#  [2] helper
# ============================================================================
def find_chrome_exe():
    cands = [
        os.path.join(os.environ.get('PROGRAMFILES', ''), 'Google/Chrome/Application/chrome.exe'),
        os.path.join(os.environ.get('PROGRAMFILES(X86)', ''), 'Google/Chrome/Application/chrome.exe'),
        os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Google/Chrome/Application/chrome.exe'),
    ]
    for p in cands:
        if p and os.path.exists(p):
            return p
    return None


def launch_kwargs(headless=False):
    kw = {
        'user_data_dir': CHROME_PROFILE,
        'headless': headless,
        'args': ['--disable-blink-features=AutomationControlled', '--start-maximized'],
        'no_viewport': True,
    }
    exe = find_chrome_exe()
    if exe:
        kw['executable_path'] = exe
    return kw


def load_prefs():
    try:
        with open(PREF_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def save_prefs(d):
    try:
        with open(PREF_FILE, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def norm_name(s):
    return re.sub(r'[\s()（）\[\]]', '', str(s or '')).lower()


def norm_bool(v):
    s = str(v or '').strip().lower()
    if s in ('yes', 'y', '1', 'true'):
        return 'yes'
    if s in ('no', 'n', '0', 'false'):
        return 'no'
    return 'any'


# ============================================================================
#  [2.5] บันทึกการใช้งาน (history)
#        เขียนทีละบรรทัดเป็น JSON ลง history.jsonl — อ่านย้อนหลังได้ทั้งหมด
#        ทุกบรรทัดมีชื่อผู้ใช้/ชื่อเครื่องติดไปด้วย เผื่อวันหน้าเอามารวมกันทั้งทีม
#        ถ้าเขียนไม่ได้ก็เงียบไป ห้ามทำให้โปรแกรมหลักพัง
# ============================================================================
SESSION_ID = datetime.now().strftime('%y%m%d%H%M%S')


def history_file():
    return os.path.join(LOCAL_LOG_DIR, 'history.jsonl')


def stats_file():
    return os.path.join(LOCAL_LOG_DIR, 'stats.json')


def set_local_log_dir(path, move=True):
    """ย้ายที่เก็บ log ในเครื่อง — ย้ายไฟล์เดิมตามไปด้วย จะได้ไม่เสียประวัติ
       คืนค่า (จำนวนไฟล์ที่ย้าย, ข้อความผิดพลาดถ้ามี)"""
    global LOCAL_LOG_DIR
    new_dir = (path or '').strip() or DATA_DIR
    old_dir = LOCAL_LOG_DIR
    moved, err = 0, ''
    try:
        os.makedirs(new_dir, exist_ok=True)
        if move and os.path.abspath(new_dir) != os.path.abspath(old_dir):
            for fn in ('history.jsonl', 'stats.json'):
                src = os.path.join(old_dir, fn)
                dst = os.path.join(new_dir, fn)
                if not os.path.exists(src):
                    continue
                if os.path.exists(dst):
                    if fn.endswith('.jsonl'):     # มีไฟล์อยู่แล้ว → ต่อท้ายไม่ให้ของเก่าหาย
                        with open(src, 'r', encoding='utf-8') as f, \
                                open(dst, 'a', encoding='utf-8') as g:
                            g.write(f.read())
                        moved += 1
                else:
                    shutil.copy2(src, dst)
                    moved += 1
    except Exception as ex:
        err = str(ex)[:150]
        new_dir = old_dir
    LOCAL_LOG_DIR = new_dir
    try:
        pr = load_prefs()
        pr['local_log_dir'] = '' if os.path.abspath(new_dir) == os.path.abspath(DATA_DIR) else new_dir
        save_prefs(pr)
    except Exception:
        pass
    return moved, err


def load_local_log_dir():
    global LOCAL_LOG_DIR
    try:
        saved = (load_prefs().get('local_log_dir') or '').strip()
        if saved and os.path.isdir(saved):
            LOCAL_LOG_DIR = saved
    except Exception:
        pass
    return LOCAL_LOG_DIR


def set_central_dir(path):
    """ตั้ง/ล้างโฟลเดอร์กลางของทีม — เก็บลง prefs ไม่ต้องแก้โค้ด"""
    global CENTRAL_LOG_DIR
    CENTRAL_LOG_DIR = (path or '').strip()
    if CENTRAL_LOG_DIR:
        try:
            os.makedirs(CENTRAL_LOG_DIR, exist_ok=True)
        except Exception:
            pass
    try:
        pr = load_prefs()
        pr['central_log_dir'] = CENTRAL_LOG_DIR
        save_prefs(pr)
    except Exception:
        pass
    return CENTRAL_LOG_DIR


def load_central_dir():
    """อ่านค่าที่เคยตั้งไว้ตอนเปิดโปรแกรม"""
    global CENTRAL_LOG_DIR
    try:
        saved = (load_prefs().get('central_log_dir') or '').strip()
        if saved:
            CENTRAL_LOG_DIR = saved
    except Exception:
        pass
    return CENTRAL_LOG_DIR


def pending_items():
    """สิ่งที่ยังตั้งค่าไม่ครบ — เอาไว้เตือนบนหน้าจอ จะได้ไม่ลืม"""
    todo = []
    if not CENTRAL_LOG_DIR:
        todo.append('ยังไม่ได้ตั้งโฟลเดอร์กลางของทีม — log ตอนนี้เก็บเฉพาะในเครื่องนี้ '
                    'ถ้าอยากรวมสถิติทั้งทีม กดปุ่ม "เลือกโฟลเดอร์" ข้างบน')
    elif not os.path.isdir(CENTRAL_LOG_DIR):
        todo.append(f'โฟลเดอร์กลางที่ตั้งไว้หายไป: {CENTRAL_LOG_DIR}')
    return todo


def _central_path():
    """ไม่สร้างโฟลเดอร์เอง — ถ้าโฟลเดอร์หาย (ไดรฟ์หลุด/ยังไม่ sync)
       ให้ข้ามไปเฉยๆ แล้วให้ pending_items() เตือนแทน
       จะได้ไม่ไปสร้างโฟลเดอร์ว่างทิ้งไว้แล้ว log กองผิดที่โดยไม่รู้ตัว"""
    if not CENTRAL_LOG_DIR or not os.path.isdir(CENTRAL_LOG_DIR):
        return None
    try:
        safe = re.sub(r'[^A-Za-z0-9._@-]', '_', f'{TR_USER}@{TR_MACHINE}')
        return os.path.join(CENTRAL_LOG_DIR, f'log_{safe}.jsonl')
    except Exception:
        return None


def log_event(_ev, **fields):
    """บันทึก 1 เหตุการณ์

    ชื่อพารามิเตอร์แรกขึ้นต้นด้วย _ เพราะฟิลด์ที่ส่งเข้ามามีชื่อ kind ด้วย
    (kind = ItemKind ของไอเทม) ถ้าตั้งชื่อซ้ำจะชนกันแล้วโปรแกรมล้ม
    """
    ev = {
        'ts': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'kind': _ev,
        'user': TR_USER,
        'machine': TR_MACHINE,
        'session': SESSION_ID,
        'app': APP_VERSION,
    }
    ev.update(fields)
    line = json.dumps(ev, ensure_ascii=False)
    for path in (history_file(), _central_path()):
        if not path:
            continue
        try:
            with open(path, 'a', encoding='utf-8') as f:
                f.write(line + '\n')
        except Exception:
            pass
    _bump_stats(_ev, fields)
    return ev


def _bump_stats(kind, fields):
    """ตัวนับรวม เอาไว้โชว์เร็วๆ โดยไม่ต้องอ่านไฟล์ทั้งก้อน"""
    try:
        st = read_stats()
        st['total_events'] = st.get('total_events', 0) + 1
        st['last_used'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        st.setdefault('first_used', st['last_used'])
        st['by_kind'] = st.get('by_kind', {})
        st['by_kind'][kind] = st['by_kind'].get(kind, 0) + 1
        day = datetime.now().strftime('%Y-%m-%d')
        st['by_day'] = st.get('by_day', {})
        st['by_day'][day] = st['by_day'].get(day, 0) + 1
        if kind == 'search_done':
            st['items_found'] = st.get('items_found', 0) + int(fields.get('found') or 0)
        if kind == 'error':
            st['errors'] = st.get('errors', 0) + 1
        with open(stats_file(), 'w', encoding='utf-8') as f:
            json.dump(st, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def read_stats():
    try:
        with open(stats_file(), 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def read_events(limit=None, kinds=None):
    """อ่าน history ย้อนหลัง (ใหม่สุดอยู่ท้ายไฟล์)"""
    out = []
    try:
        with open(history_file(), 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    ev = json.loads(line)
                except Exception:
                    continue
                if kinds and ev.get('kind') not in kinds:
                    continue
                out.append(ev)
    except Exception:
        return []
    return out[-limit:] if limit else out


def trim_history():
    """ตัด log เก่าทิ้งถ้ายาวเกิน LOG_MAX_EVENTS"""
    try:
        hf = history_file()
        if not os.path.exists(hf):
            return
        with open(hf, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        if len(lines) <= LOG_MAX_EVENTS:
            return
        with open(hf, 'w', encoding='utf-8') as f:
            f.writelines(lines[-LOG_MAX_EVENTS:])
    except Exception:
        pass


# ============================================================================
#  [3] Template Excel
# ============================================================================
def build_template(path):
    notes = [
        'ItemKind = รหัสไอเท็ม (ช่องค้นหาหลัก)   |   ItemName = ใส่เพื่อกรองชื่อซ้ำอีกชั้น (ไม่บังคับ)',
        'duration : เว้นว่าง = เฉพาะไอเท็มถาวร (ระยะเวลา = 0)  |  ใส่ตัวเลข = จำนวนวันนั้น (เช่น 15)  |  any = ไม่กรอง',
        'trade : Any / Yes / No        qty : เว้นว่าง = ไม่กรอง หรือใส่ค่าที่ต้องตรงเป๊ะ',
        'กรอกรายการตั้งแต่แถวที่ 5 ลงมา',
    ]
    if XLSX_OK and path.lower().endswith(('.xlsx', '.xlsm')):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Items'
        for r, line in enumerate(notes, 1):
            c = ws.cell(row=r, column=1, value=line)
            c.font = Font(italic=True, color='8B949E', size=9)
            c.fill = PatternFill('solid', fgColor='21262D')
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(TMPL_HEADERS))
        blue = PatternFill('solid', fgColor='1F6FEB')
        orange = PatternFill('solid', fgColor='D29922')
        for col, h in enumerate(TMPL_HEADERS, 1):
            cell = ws.cell(row=5, column=col, value=h)
            search_col = h in ('ItemKind', 'ItemName')
            cell.fill = blue if search_col else orange
            cell.font = Font(color='FFFFFF' if search_col else '0D1117', bold=True)
            cell.alignment = Alignment(horizontal='center')
            ws.column_dimensions[cell.column_letter].width = [14, 30, 12, 10, 10][col - 1]
        ws.append(['40852', '', '', 'No', '1'])
        ws.append(['40852', '', '15', 'Yes', ''])
        ws.append(['40852', '', 'any', 'Any', ''])
        ws.add_data_validation(DataValidation(
            type='list', formula1='"Any,Yes,No"', allow_blank=True, sqref='D6:D1000'))
        wb.save(path)
        return
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(TMPL_HEADERS)
        w.writerow(['40852', '', '', 'No', '1'])


def read_template(path):
    """อ่านไฟล์ CSV แบบ template ง่ายๆ (ไฟล์ Excel ใช้หน้าต่างนำเข้าแทน)"""
    rows = []
    with open(path, 'r', encoding='utf-8-sig', errors='replace') as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames or []
        cmap = {k: guess_column(headers, k) for k in ('kind', 'name', 'dur', 'trade', 'qty')}
        for d in reader:
            cells = [str(d.get(h, '') or '') for h in headers]
            row = cells_to_row(cells, cmap)
            if row:
                rows.append(row)
    return rows


# ============================================================================
#  [3.5] อ่านไฟล์ Excel ต้นฉบับ (ชีทเยอะ หัวตารางอยู่ลึก และมีหลายบล็อกต่อชีท)
# ============================================================================
def num_str(v):
    """แปลงค่าจากเซลล์เป็นข้อความ — ตัด .0 ที่ Excel ติดมากับตัวเลข"""
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    s = str(v).strip()
    m = re.fullmatch(r'(\d+)\.0+', s)
    return m.group(1) if m else s


def dur_from_text(v):
    """กติกาเดียวกับเครื่องมือเดิมเป๊ะ

         เว้นว่าง   ->  ''     = เฉพาะไอเท็มถาวร (บนเว็บช่องระยะเวลาไอเท็ม = 0)
         'ถาวร'     ->  ''     = ความหมายเดียวกัน (ชีทต้นฉบับเขียนคำนี้)
         'any'      ->  'any'  = ไม่กรองระยะเวลา
         ตัวเลข     ->  ตัวเลขนั้น เช่น '15'
    """
    s = num_str(v).strip()
    if not s:
        return ''
    if 'ถาวร' in s or 'permanent' in s.lower():
        return ''
    if s.lower() == 'any':
        return 'any'
    m = re.search(r'(\d+)', s)
    return m.group(1) if m else 'any'


def _norm_h(s):
    return re.sub(r'[\s_\-]', '', str(s or '')).strip().lower()


def _is_marker_cell(c):
    """เซลล์นี้เป็น 'ชื่อคอลัมน์' จริงไหม — ต้องสั้นและขึ้นต้นด้วยคำหลัก
       (กันบรรทัดคำอธิบายยาวๆ ที่บังเอิญมีคำว่า ItemKind อยู่ในประโยค)"""
    h = _norm_h(c)
    if not h or len(h) > 24:
        return False
    return any(h == m or h.startswith(m) for m in HEADER_MARKERS)


def is_header_row(cells):
    if sum(1 for c in cells if _norm_h(c)) < 2:      # หัวตารางต้องมีอย่างน้อย 2 ช่อง
        return False
    return any(_is_marker_cell(c) for c in cells)


def guess_column(headers, key):
    """เดาว่าคอลัมน์ไหนคือ key ('kind'/'name'/…) คืน index หรือ -1 ถ้าไม่เจอ"""
    hs = [_norm_h(h) for h in headers]
    for want in ALIAS.get(key, []):
        w = _norm_h(want)
        for i, h in enumerate(hs):
            if h == w:
                return i
    for want in ALIAS.get(key, []):
        w = _norm_h(want)
        for i, h in enumerate(hs):
            if h and (w in h or h in w):
                return i
    return -1


def cells_to_row(cells, cmap):
    """แปลงข้อมูล 1 แถว + การจับคู่คอลัมน์ → เงื่อนไขค้นหา 1 รายการ"""
    def get(k):
        i = cmap.get(k, -1)
        return str(cells[i]).strip() if 0 <= i < len(cells) else ''

    kind = num_str(get('kind'))
    name = get('name')
    if cmap.get('kind', -1) >= 0:
        if not re.fullmatch(r'\d+', kind):      # ตัดแถวหมายเหตุ/ยอดรวมทิ้ง
            return None
    elif not name:
        return None
    return {
        'kind': kind,
        'name': name,
        'dur': dur_from_text(get('dur')) if cmap.get('dur', -1) >= 0 else 'any',
        'trade': norm_bool(get('trade')) if cmap.get('trade', -1) >= 0 else 'any',
        'qty': get('qty') if cmap.get('qty', -1) >= 0 else '',
    }


def open_workbook(path):
    """เปิดไฟล์ Excel ครั้งเดียวแล้วใช้ซ้ำ — ไฟล์ใหญ่การเปิดใหม่ทุกครั้งช้ามาก"""
    return openpyxl.load_workbook(path, read_only=True, data_only=True)


def scan_sheets_wb(wb, progress=None):
    """คืน [(ชื่อชีท, จำนวนบล็อกตารางไอเทมที่เจอ)]"""
    out = []
    names = wb.sheetnames
    for k, name in enumerate(names):
        n = 0
        try:
            for row in wb[name].iter_rows(min_row=1, max_row=SCAN_ROWS, values_only=True):
                if row and is_header_row(row):
                    n += 1
        except Exception:
            n = 0
        out.append((name, n))
        if progress:
            progress(k + 1, len(names), name)
    return out


def scan_sheets(path):
    wb = open_workbook(path)
    try:
        return scan_sheets_wb(wb)
    finally:
        wb.close()


def parse_sheet_wb(wb, sheet):
    """อ่านชีทเดียว คืน (headers, blocks)
       blocks = [{'row': แถวหัวตาราง, 'headers': [...], 'data': [[cell, …], …]}]"""
    rows = [list(r) if r else []
            for r in wb[sheet].iter_rows(min_row=1, max_row=SCAN_ROWS, values_only=True)]

    blocks, i = [], 0
    while i < len(rows):
        if rows[i] and is_header_row(rows[i]):
            headers = [num_str(c) for c in rows[i]]
            data, j = [], i + 1
            while j < len(rows) and len(data) < 500:
                r = rows[j]
                if not r or is_header_row(r) or not any(num_str(c) for c in r):
                    break
                data.append([num_str(c) for c in r])
                j += 1
            blocks.append({'row': i + 1, 'headers': headers, 'data': data})
            i = max(j, i + 1)
        else:
            i += 1
    return (blocks[0]['headers'] if blocks else []), blocks


def parse_sheet(path, sheet):
    wb = open_workbook(path)
    try:
        return parse_sheet_wb(wb, sheet)
    finally:
        wb.close()


def blocks_to_rows(blocks, cmap):
    out = []
    for b in blocks:
        for cells in b['data']:
            row = cells_to_row(cells, cmap)
            if row:
                out.append(row)
    return out


# ============================================================================
#  [4] ตรรกะ deep check (ยกมาจากเครื่องมือเดิม)
# ============================================================================
def has_deep(c):
    return c.get('trade', 'any') != 'any' or str(c.get('qty', '')).strip() != '' \
        or str(c.get('dur', 'any')).strip().lower() != 'any'


def duration_only_numeric(c):
    """กรองแค่ระยะเวลาเป็นตัวเลข → ใช้ฟิลเตอร์บนหน้า list ได้ ไม่ต้องเปิดรายละเอียด"""
    return c.get('trade', 'any') == 'any' and not str(c.get('qty', '')).strip() \
        and re.fullmatch(r'\d+', str(c.get('dur', '')).strip() or '')


def match_deep(detail, crit):
    notes = []
    want = str(crit.get('dur', 'any')).strip().lower()
    actual = str(detail.get('duration') or '').strip()
    if want != 'any':
        if want == '':
            if actual not in ('', '0'):
                return False, f'ไม่ใช่ถาวร (dur={actual})'
            notes.append('ถาวร')
        elif actual != want:
            return False, f'dur {actual or "-"} ≠ {want}'
        else:
            notes.append(f'{want} วัน')
    elif actual:
        notes.append(f'{actual} วัน')

    trade = detail.get('trade')
    if crit.get('trade', 'any') != 'any':
        if trade is None:
            return False, 'อ่านค่าแลกเปลี่ยนไม่ได้'
        if trade != (crit['trade'] == 'yes'):
            return False, 'trade=' + ('Y' if trade else 'N')
    if trade is not None:
        notes.append('trade=' + ('Y' if trade else 'N'))

    q = str(crit.get('qty', '')).strip()
    if q:
        if str(detail.get('qty') or '').strip() != q:
            return False, f'qty {detail.get("qty") or "-"} ≠ {q}'
        notes.append('qty=' + q)
    return True, ' · '.join(notes)


# ============================================================================
#  [5] JS ที่ยิงเข้าไปอ่านหน้าเว็บ
# ============================================================================
JS_READ_ROWS = """
(col) => {
  const out = [];
  document.querySelectorAll('table tbody tr').forEach(tr => {
    const td = [...tr.querySelectorAll('td')];
    if (td.length < 4) return;
    const id = (td[col.id].innerText || '').trim();
    if (!/^\\d+$/.test(id)) return;
    out.push({ id,
      name: (td[col.name].innerText || '').trim(),
      type: (td[col.type].innerText || '').trim(),
      kind: (td[col.kind].innerText || '').trim() });
  });
  return out;
}"""

JS_TOTAL_PAGES = """
() => {
  const i = document.querySelector('input[type="number"]');
  if (i && i.max) return parseInt(i.max, 10) || 1;
  const m = (document.body.innerText || '').match(/of\\s+(\\d+)/);
  return m ? parseInt(m[1], 10) : 1;
}"""

JS_READ_DETAIL = """
(sel) => {
  const v = s => { const e = document.querySelector(s); return e ? (e.value || '').trim() : null; };
  function near(el, n = 5) {
    let c = el;
    for (let i = 0; i < n && c.parentElement; i++) {
      c = c.parentElement;
      const t = (c.innerText || '').trim();
      if (t && t.length < 160) return t;
    }
    return '';
  }
  let trade = null;
  const boxes = [...document.querySelectorAll('input[type="checkbox"]'),
                 ...document.querySelectorAll('[role="checkbox"]'),
                 ...document.querySelectorAll('[role="switch"]')];
  for (const cb of boxes) {
    if (near(cb).includes(sel.tradeLabel)) {
      trade = cb.type === 'checkbox' ? !!cb.checked
        : (cb.getAttribute('data-state') === 'checked' || cb.getAttribute('aria-checked') === 'true');
      break;
    }
  }
  return { name: v(sel.name), kind: v(sel.kind), price: v(sel.price),
           duration: v(sel.duration), qty: v(sel.qty), trade };
}"""

JS_CLICK_ROW = """
(id) => {
  for (const tr of document.querySelectorAll('table tbody tr')) {
    const td = tr.querySelectorAll('td');
    if (td.length && (td[0].innerText || '').trim() === String(id)) {
      const a = tr.querySelector('a');
      if (a) { a.click(); return true; }
    }
  }
  return false;
}"""

# ---------------------------------------------------------------------------
#  JS สำหรับหน้า "สร้างไอเทม"
#  เว็บเป็น React — กรอกด้วย .value เฉยๆ ไม่พอ ต้องใช้ native setter
#  แล้วยิง event input/change เอง ไม่งั้น React ไม่รู้ว่าค่าเปลี่ยน
#  (วิธีเดียวกับเครื่องมือเดิม)
# ---------------------------------------------------------------------------
JS_SET_VALUE = """
([sel, val]) => {
  const el = document.querySelector(sel);
  if (!el) return 'ไม่เจอช่อง';
  el.scrollIntoView({block:'center'});
  const proto = el.tagName === 'TEXTAREA'
      ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
  const d = Object.getOwnPropertyDescriptor(proto, 'value');
  if (d && d.set) { d.set.call(el, String(val)); } else { el.value = String(val); }
  el.dispatchEvent(new Event('input',  {bubbles:true}));
  el.dispatchEvent(new Event('change', {bubbles:true}));
  return 'ok';
}"""

# กรอกโดยอ้างข้อความที่อยู่ข้างๆ — ใช้เป็นตัวสำรองเวลา id เปลี่ยน
JS_SET_BY_LABEL = """
([lbl, val]) => {
  const vis = el => { if(!el) return false; const s=getComputedStyle(el);
    if(s.display==='none'||s.visibility==='hidden') return false;
    const r=el.getBoundingClientRect(); return r.width>0 && r.height>0; };
  const hit = t => { t=(t||'').trim();
    return t===lbl || t.replace(/\\s*\\*\\s*$/,'')===lbl ||
           (t.length<=lbl.length+3 && t.indexOf(lbl)===0); };
  const w = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
  let node, target=null;
  while ((node = w.nextNode()) && !target) {
    if (!hit(node.textContent)) continue;
    let c = node.parentElement;
    for (let i=0; i<5 && c && !target; i++) {
      for (const el of c.querySelectorAll('input,textarea')) {
        if (vis(el) && el.type!=='checkbox' && el.type!=='file') { target = el; break; }
      }
      c = c.parentElement;
    }
  }
  if (!target) return 'ไม่เจอช่อง';
  target.scrollIntoView({block:'center'});
  const proto = target.tagName === 'TEXTAREA'
      ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
  Object.getOwnPropertyDescriptor(proto,'value').set.call(target, String(val));
  target.dispatchEvent(new Event('input',  {bubbles:true}));
  target.dispatchEvent(new Event('change', {bubbles:true}));
  return 'ok';
}"""

# ติ๊ก/ปลดติ๊ก checkbox หรือสวิตช์ ที่อยู่ใกล้ข้อความที่ระบุ
JS_SET_SWITCH = """
([lbl, on]) => {
  const nz = t => (t||'').trim();
  const nodes = [...document.querySelectorAll('*')].filter(
    e => e.children.length === 0 && (nz(e.textContent) === lbl ||
         nz(e.textContent).indexOf(lbl) === 0));
  for (const el of nodes) {
    let n = el.parentElement;
    for (let i=0; i<6 && n; i++) {
      const cb = n.querySelector('input[type="checkbox"],[role="checkbox"],[role="switch"]');
      if (cb) {
        const cur = (cb.checked !== undefined)
            ? cb.checked : (cb.getAttribute('aria-checked') === 'true');
        if (cur !== on) cb.click();
        return 'ok';
      }
      n = n.parentElement;
    }
  }
  return 'ไม่เจอสวิตช์';
}"""

# เลือกประเภทไอเทม (game_item_type) — รองรับทั้ง <select> และ combobox ของ React
JS_PICK_TYPE = """
([val, ph]) => {
  for (const s of document.querySelectorAll('select')) {
    const opts = [...s.options];
    let i = opts.findIndex(o => o.text.trim().toLowerCase() === val.toLowerCase());
    if (i < 0) i = opts.findIndex(o => o.text.trim().toLowerCase().includes(val.toLowerCase()));
    if (i >= 0) {
      const d = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value');
      d.set.call(s, opts[i].value);
      s.dispatchEvent(new Event('input',  {bubbles:true}));
      s.dispatchEvent(new Event('change', {bubbles:true}));
      return 'ok';
    }
  }
  return 'combobox';
}"""

# อ่านค่าที่อยู่ในฟอร์มตอนนี้กลับมา — ใช้ตรวจว่ากรอกเข้าไปจริงไหม
JS_READ_FORM = """
(sel) => {
  const v = s => { const e = document.querySelector(s); return e ? e.value : null; };
  return { name: v(sel.name), kind: v(sel.kind), price: v(sel.price),
           duration: v(sel.duration), qty: v(sel.qty), mail: v(sel.mail),
           desc: v(sel.desc) };
}"""


# ============================================================================
#  [5.4] อ่านชีทต้นฉบับแบบเดียวกับเครื่องมือเดิม (parse_master_rows)
#        ไม่ต้องกรอก template เอง — ยกกฎมาจาก tr_studio.py ทั้งดุ้น
#
#        anchor  = คอลัมน์ที่หัวตารางเขียนว่า fdItemNum (หรือ Item Kind)
#        name    = Name / Display Name / Item Name  ที่อยู่ในช่วงของ anchor นั้น
#        qty     = เลขที่อยู่หน้าคำว่า "ชิ้น" ในชื่อ   เช่น "กล่องอิลลิเฟียร์ 5 ชิ้น" -> 5
#        ชื่อโชว์ = ตัด " X ชิ้น" ท้ายออก                 -> "กล่องอิลลิเฟียร์"
#        dur     = เลขในคอลัมน์ระยะเวลา เฉพาะตอนมีคำว่า "วัน" ; "ถาวร"/ว่าง -> '' (ถาวร)
#        trade   = คอลัมน์ Itemmove (Yes/No, ว่าง = No) ; ไม่มีคอลัมน์เลย -> 'any'
#        หนึ่งแถวมีได้หลายตาราง (fdItemNum หลายคอลัมน์วางเรียงกัน) และตัด ID ซ้ำให้
# ============================================================================
MASTER_TRADE_YES = ('yes', 'y', 'true', 'ได้', 'แลกเปลี่ยนได้')


def read_sheet_rows(wb, sheet):
    return [list(r) if r else []
            for r in wb[sheet].iter_rows(min_row=1, max_row=SCAN_ROWS, values_only=True)]


MAX_SANE_NAME = 70          # ชื่อไอเทมยาวเกินนี้ ปกติจะเป็นข้อความโน้ตในชีท ไม่ใช่ชื่อไอเทม
#   วัดจากไฟล์จริงแล้ว ชื่อที่ยาวที่สุด (ประกอบคำต่อท้าย+วัน แล้ว) = 52 ตัวอักษร
#   ตั้ง 70 เผื่อไว้ ของจริงจะไม่โดนเตือน แต่ประโยคโน้ตในชีทจะโดน


def kind_mode_len(rows):
    """ความยาวเลข ItemKind ที่พบบ่อยที่สุดในชุดนี้ — ใช้เป็นไม้บรรทัดของชุดนั้นเอง
    ไม่ได้ตั้งกฎตายตัวว่าต้องกี่หลัก เพราะแต่ละไฟล์อาจไม่เหมือนกัน"""
    counts = {}
    for r in rows:
        k = str(r.get('kind') or '')
        if k:
            counts[len(k)] = counts.get(len(k), 0) + 1
    if not counts:
        return 0
    return max(counts.items(), key=lambda kv: (kv[1], kv[0]))[0]


def suspect_reasons(r, mode_len=0):
    """บอกว่าแถวนี้ดู "แปลกปลอม" ตรงไหน — บอกเฉยๆ ไม่ตัดทิ้ง
    ชีทบางใบเขียนโน้ตปนมาในคอลัมน์ตาราง เลยโผล่มาเป็นไอเทมได้
    แต่บางทีก็เป็นของจริง เลยให้คนตัดสินเองเป็นเคสๆ ไป"""
    out = []
    k = str(r.get('kind') or '')
    if mode_len and k and len(k) <= mode_len - 2:
        out.append('เลข ItemKind สั้นกว่าตัวอื่นในชุดนี้มาก (%d หลัก · ส่วนใหญ่ %d หลัก)'
                   % (len(k), mode_len))
    nm = str(r.get('raw') or r.get('name') or r.get('disp') or '')
    if len(nm) > MAX_SANE_NAME:
        out.append('ชื่อยาวผิดปกติ (%d ตัวอักษร) อาจเป็นข้อความโน้ตในชีท' % len(nm))
    return out


def mark_suspects(rows):
    """ติดป้าย warn ให้ทุกแถว แล้วคืนจำนวนแถวที่น่าสงสัย (ไม่ลบแถวไหนทิ้งเลย)"""
    m = kind_mode_len(rows)
    n = 0
    for r in rows:
        w = suspect_reasons(r, m)
        r['warn'] = w
        if w:
            n += 1
    return n


def split_name_qty(raw_name):
    """แยก "ชื่อ" กับ "จำนวนชิ้น" ออกจากกัน — กฎเดียวกับเครื่องมือเดิมเป๊ะ

    ตัด "N ชิ้น" ตัวสุดท้ายออกจากชื่อ แล้วต่อส่วนหน้า+ส่วนหลังเข้าด้วยกัน
    -> จำนวนชิ้น "ไม่" ติดไปในชื่อ แต่ ชม./อย่างอื่น ยังอยู่ครบ
       'กล่อง 2 ชั่วโมง 3 ชิ้น' -> ชื่อ 'กล่อง 2 ชั่วโมง' · จำนวน 3
    ชื่อที่ไม่มีคำว่า 'ชิ้น' -> คงชื่อเดิม จำนวน 1
    คืนค่า (ชื่อ, จำนวน, มีคำว่าชิ้นไหม)
    """
    s = (raw_name or '').strip()
    mqs = list(re.finditer(r'(\d+)\s*ชิ้น', s))
    if not mqs:
        return s, '1', False
    m = mqs[-1]
    base = (s[:m.start()] + s[m.end():]).strip()
    base = re.sub(r'\s{2,}', ' ', base)
    return base, m.group(1), True


def make_item_name(cname, suffix='', dur=''):
    """ประกอบชื่อที่จะกรอกลงเว็บ: <ชื่อ> <คำต่อท้าย> (<X> วัน)

    กฎเดิม: คำต่อท้ายกับ '(X วัน)' เติมเฉพาะชื่อที่เคยมี 'ชิ้น' เท่านั้น
    (ตัวเรียกเป็นคนตัดสินใจ — ฟังก์ชันนี้ต่อให้ตามที่ส่งมา)
    """
    s = (cname or '').strip()
    suffix = (suffix or '').strip()
    dur = str(dur or '').strip()
    if suffix:
        s = (s + ' ' + suffix).strip()
    if dur:
        s = (s + ' (%s วัน)' % dur).strip()
    return s


def row_to_item(r, type_name=DEFAULT_TYPE, suffix=DEFAULT_SUFFIX, price='', mail=''):
    """แปลงแถวที่อ่านจากชีท -> ข้อมูลไอเทมสำหรับหน้า create"""
    cname = r.get('cname') or r.get('disp') or ''
    if r.get('has_qty'):
        name = make_item_name(cname, suffix, r.get('dur', ''))
    else:
        name = cname          # ชื่อไม่มี 'ชิ้น' -> ไม่เติมอะไรเลย
    return {
        'kind': r.get('kind', ''),
        'name': name,
        'cname': cname,
        'desc': r.get('desc', ''),
        'type': type_name,
        'price': str(price or ''),
        'dur': r.get('dur', ''),
        'mail': str(mail or ''),
        'qty': r.get('amount', '1'),
        'trade': r.get('trade', 'any') == 'yes',
        'web': True,
    }


def parse_master_rows(rows, seen=None):
    out = []
    if seen is None:
        seen = set()
    tables = []
    prev = []

    for row in rows:
        cells = [num_str(c) for c in row]
        low = [c.lower() for c in cells]

        anchors = [j for j, c in enumerate(low) if 'fditemnum' in c]
        if not anchors:
            anchors = [j for j, c in enumerate(low) if c in ('item kind', 'itemkind')]

        if anchors:
            top = [str(c or '').strip().lower() for c in prev]
            tables = []
            for idx, a in enumerate(anchors):
                end_c = anchors[idx + 1] if idx + 1 < len(anchors) else max(len(low), a + 12)
                span = list(range(a, min(end_c, len(low))))

                def findin(preds, _span=span, _low=low):
                    for j in _span:
                        h = _low[j]
                        if any(pr(h) for pr in preds):
                            return j
                    return None

                name_col = findin([lambda h: h == 'name', lambda h: 'display name' in h,
                                   lambda h: h == 'item name', lambda h: h == 'itemname'])
                desc_col = findin([lambda h: 'item des' in h, lambda h: h == 'des',
                                   lambda h: 'description' in h, lambda h: 'รายละเอียด' in h,
                                   lambda h: 'item dec' in h])
                dur_col = findin([lambda h: 'ระยะเวลา' in h, lambda h: 'ของขวัญ' in h,
                                  lambda h: 'duration' in h])
                move_col = findin([lambda h: 'itemmove' in h])
                if move_col is None:      # บางไฟล์เขียน Itemmove ไว้แถวบนของหัวตาราง
                    for j in span:
                        if j < len(top) and 'itemmove' in top[j]:
                            move_col = j
                            break
                tables.append({'kind': a, 'name': name_col, 'dur': dur_col,
                               'move': move_col, 'desc': desc_col})
            prev = cells
            continue

        prev = cells
        if not tables:
            continue

        for t in tables:
            a = t['kind']
            if a >= len(cells):
                continue
            kind = num_str(cells[a])
            if not re.fullmatch(r'\d+', kind):
                continue
            nc = t['name'] if t['name'] is not None else (a + 4)
            name = cells[nc] if nc < len(cells) else ''
            # ต้องมีตัวอักษรจริงในชื่อ ไม่ใช่ตัวเลขล้วน (กันแถวยอดรวม)
            if not name.strip() or not re.search(r'[^\d.,\s]', name):
                continue
            if kind in seen:
                continue
            seen.add(kind)

            m = re.search(r'(\d+)\s*ชิ้น', name)
            qty = m.group(1) if m else ''
            disp = re.sub(r'\s*\d+\s*ชิ้น.*$', '', name).strip() or name.strip()

            dc = t['dur'] if t['dur'] is not None else (a + 5)
            durc = cells[dc] if dc < len(cells) else ''
            md = re.search(r'(\d+)', durc) if 'วัน' in durc else None
            dur = md.group(1) if md else ''

            if t['move'] is not None:
                mc = cells[t['move']] if t['move'] < len(cells) else ''
                trade = 'yes' if mc.strip().lower() in MASTER_TRADE_YES else 'no'
            else:
                trade = 'any'

            dsc = ''
            if t.get('desc') is not None and t['desc'] < len(cells):
                dsc = cells[t['desc']].strip()

            cname, amount, has_qty = split_name_qty(name)

            out.append({'kind': kind, 'name': '', 'disp': disp,
                        'dur': dur, 'trade': trade, 'qty': qty,
                        'has_move': t['move'] is not None,
                        # ---- ใช้ตอน "สร้างไอเทม" ----
                        'raw': name.strip(), 'cname': cname,
                        'amount': amount, 'has_qty': has_qty, 'desc': dsc})
    return out


# ============================================================================
#  [5.5] หน้าต่างนำเข้า Excel — เลือกชีท แล้วอ่านให้อัตโนมัติ
# ============================================================================
class ImportDialog:
    def __init__(self, parent, path):
        self.path = path
        self.result = None
        self.sheet_name = ''
        self.sheet_names = []
        self.sheets = []
        self.rows = []
        self._pending = None

        self.top = tk.Toplevel(parent)
        self.top.title('นำเข้าจาก Excel')
        self.top.configure(bg=C['bg'])
        self.top.geometry('1010x640')
        self.top.transient(parent)
        self.top.grab_set()

        head = tk.Frame(self.top, bg=C['card'], height=56)
        head.pack(fill='x')
        head.pack_propagate(False)
        tk.Label(head, text='นำเข้าจาก Excel', bg=C['card'], fg=C['fg'],
                 font=('Segoe UI', 13, 'bold')).pack(side='left', padx=18)
        tk.Label(head, text=os.path.basename(path), bg=C['card'], fg=C['dim'],
                 font=('Segoe UI', 9)).pack(side='left')

        foot = tk.Frame(self.top, bg=C['card'], height=62)
        foot.pack(fill='x', side='bottom')
        foot.pack_propagate(False)
        self.count_lbl = tk.Label(foot, text='', bg=C['card'], fg=C['dim'], font=('Segoe UI', 10))
        self.count_lbl.pack(side='left', padx=18)
        self.ok_btn = tk.Button(foot, text='ใช้ข้อมูลนี้', bg=C['accent'], fg='white', bd=0,
                                font=('Segoe UI', 10, 'bold'), cursor='hand2',
                                activebackground='#3d55cf', activeforeground='white',
                                state='disabled', command=self._ok)
        self.ok_btn.pack(side='right', padx=18, pady=12, ipadx=22, ipady=5)
        tk.Button(foot, text='ยกเลิก', bg=C['input'], fg=C['fg'], bd=0,
                  font=('Segoe UI', 10), cursor='hand2', activebackground=C['line'],
                  command=self._cancel).pack(side='right', pady=12, ipadx=16, ipady=5)

        body = tk.Frame(self.top, bg=C['bg'])
        body.pack(fill='both', expand=True, padx=14, pady=12)

        # ---- ซ้าย : รายชื่อชีท ----
        left = tk.Frame(body, bg=C['bg'], width=280)
        left.pack(side='left', fill='y')
        left.pack_propagate(False)
        tk.Label(left, text='เลือกชีท', bg=C['bg'], fg=C['dim'],
                 font=('Segoe UI', 9, 'bold')).pack(anchor='w')
        lbwrap = tk.Frame(left, bg=C['bg'])
        lbwrap.pack(fill='both', expand=True, pady=(5, 0))
        # selectmode='multiple' = คลิกทีละชีทเพื่อเลือก/ยกเลิก เลือกหลายชีทได้โดยไม่ต้องกด Ctrl
        self.lb = tk.Listbox(lbwrap, bg=C['input'], fg=C['fg'], bd=0,
                             highlightthickness=1, highlightbackground=C['line'],
                             selectbackground=C['accent'], selectforeground='white',
                             font=('Segoe UI', 9), activestyle='none',
                             selectmode='multiple')
        sb = ttk.Scrollbar(lbwrap, orient='vertical', command=self.lb.yview)
        self.lb.configure(yscrollcommand=sb.set)
        self.lb.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')
        self.lb.bind('<<ListboxSelect>>', lambda e: self._on_sheet())

        self.sel_lbl = tk.Label(left, text='', bg=C['bg'], fg=C['fg'],
                                font=('Segoe UI', 9, 'bold'), anchor='w')
        self.sel_lbl.pack(fill='x', pady=(8, 2))
        btnf = tk.Frame(left, bg=C['bg'])
        btnf.pack(fill='x')
        for txt, cmd in (('เลือกชีทที่มีตารางทั้งหมด', self._pick_all),
                         ('ล้างที่เลือก', self._pick_none)):
            tk.Button(btnf, text=txt, bg=C['input'], fg=C['fg'], bd=0,
                      font=('Segoe UI', 9), cursor='hand2', activebackground=C['line'],
                      command=cmd).pack(side='left', padx=(0, 6), ipadx=6, ipady=3)
        tk.Label(left, text='คลิกชีทเพื่อเลือก คลิกซ้ำเพื่อเอาออก — เลือกได้หลายชีทพร้อมกัน '
                            'ID ที่ซ้ำกันข้ามชีทจะถูกตัดให้เหลือตัวเดียว',
                 bg=C['bg'], fg=C['dim'], font=('Segoe UI', 8), anchor='w',
                 justify='left', wraplength=265).pack(fill='x', pady=(8, 0))

        # ---- ขวา ----
        right = tk.Frame(body, bg=C['bg'])
        right.pack(side='left', fill='both', expand=True, padx=(14, 0))

        self.info = tk.Label(right, text='กำลังสแกนไฟล์…', bg=C['bg'], fg=C['dim'],
                             font=('Segoe UI', 9), anchor='w', justify='left')
        self.info.pack(fill='x')

        tk.Label(right, text='อ่านให้อัตโนมัติแบบเดียวกับเครื่องมือเดิม — '
                             'qty ดึงจากเลขหน้าคำว่า “ชิ้น” ในชื่อ · ระยะเวลาอ่านจากคอลัมน์ที่มีคำว่า “วัน” · '
                             'แลกเปลี่ยนอ่านจากคอลัมน์ Itemmove (ไม่มีคอลัมน์ = ไม่กรอง)',
                 bg=C['bg'], fg=C['dim'], font=('Segoe UI', 8), wraplength=690,
                 justify='left').pack(anchor='w', pady=(6, 0))

        self.warn_lbl = tk.Label(right, text='', bg=C['bg'], fg=C['warn'],
                                 font=('Segoe UI', 9), anchor='w', justify='left',
                                 wraplength=690)
        self.warn_lbl.pack(fill='x', pady=(6, 0))

        tk.Label(right, text='ข้อมูลที่จะนำเข้า', bg=C['bg'], fg=C['dim'],
                 font=('Segoe UI', 9, 'bold')).pack(anchor='w', pady=(12, 4))
        pv = tk.Frame(right, bg=C['bg'])
        pv.pack(fill='both', expand=True)
        cols = ('w', 'kind', 'name', 'dur', 'trade', 'qty')
        self.tree = ttk.Treeview(pv, columns=cols, show='headings', style='TR.Treeview', height=13)
        for c, t, w in (('w', '', 26), ('kind', 'ItemKind', 90), ('name', 'ชื่อไอเทม', 300),
                        ('dur', 'ระยะเวลา', 90), ('trade', 'แลกเปลี่ยน', 90),
                        ('qty', 'จำนวน', 70)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor='w')
        self.tree.tag_configure('warn', background='#3a3018', foreground='#e3b341')
        self.tree.bind('<<TreeviewSelect>>', lambda e: self._why())
        tsb = ttk.Scrollbar(pv, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=tsb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        tsb.pack(side='right', fill='y')

        self.top.protocol('WM_DELETE_WINDOW', self._cancel)
        self.wb = None
        self.cache = {}
        self.lock = threading.Lock()
        self.q = queue.Queue()
        self.inflight = set()
        threading.Thread(target=self._scan_worker, daemon=True).start()
        self.top.after(80, self._pump)
        parent.wait_window(self.top)

    # ---------- รับผลจากเธรดเบื้องหลัง (บนเธรดหลักเท่านั้น) ----------
    def _pump(self):
        try:
            while True:
                msg = self.q.get_nowait()
                kind = msg[0]
                if kind == 'info':
                    self.info.config(text=msg[1])
                elif kind == 'sheets':
                    self.sheets = msg[1]
                    self._fill_sheets()
                elif kind == 'rows':
                    key, rows, note = msg[1], msg[2], msg[3]
                    self.cache[key] = (rows, note)
                    self.inflight.discard(key)
                    if self._key() == key:
                        self._show(rows, note)
        except queue.Empty:
            pass
        except Exception:
            pass
        try:
            self.top.after(80, self._pump)
        except Exception:
            pass

    # ---------- สแกนรายชื่อชีท ----------
    def _scan_worker(self):
        try:
            with self.lock:
                self.q.put(('info', 'กำลังเปิดไฟล์…'))
                self.wb = open_workbook(self.path)
                sheets = scan_sheets_wb(
                    self.wb,
                    progress=lambda k, n, name: self.q.put(
                        ('info', f'กำลังสแกน… {k}/{n}   {name}')))
        except Exception as ex:
            self.q.put(('info', 'อ่านไฟล์ไม่ได้: ' + str(ex)))
            return
        self.q.put(('sheets', sheets))

    def _fill_sheets(self):
        self.lb.delete(0, tk.END)
        first_hit = None
        for i, (name, n) in enumerate(self.sheets):
            self.lb.insert(tk.END, (f'★ ({n})  ' if n else '     ') + name)
            if n and first_hit is None:
                first_hit = i
        hits = sum(1 for _, n in self.sheets if n)
        self.info.config(text=f'ไฟล์นี้มี {len(self.sheets)} ชีท · พบตารางไอเทมใน {hits} ชีท')
        if first_hit is not None:
            self.lb.selection_set(first_hit)
            self.lb.see(first_hit)
            self._on_sheet()

    # ---------- เลือกชีท (เลือกได้หลายชีทพร้อมกัน) ----------
    def _selected(self):
        """ชื่อชีทที่ถูกเลือกอยู่ เรียงตามลำดับในไฟล์"""
        return [self.sheets[i][0] for i in self.lb.curselection()
                if 0 <= i < len(self.sheets)]

    def _key(self):
        return '\x00'.join(self._selected())

    def _pick_all(self):
        self.lb.selection_clear(0, tk.END)
        for i, (_, n) in enumerate(self.sheets):
            if n:
                self.lb.selection_set(i)
        self._on_sheet()

    def _pick_none(self):
        self.lb.selection_clear(0, tk.END)
        self._on_sheet()

    def _on_sheet(self):
        """คลิกรัวๆ หลายชีทติดกัน ไม่ต้องอ่านไฟล์ทุกคลิก — รอให้หยุดคลิกก่อน"""
        names = self._selected()
        self.sheet_names = names
        self.sheet_name = names[0] if names else ''
        self.sel_lbl.config(text=(f'เลือกไว้ {len(names)} ชีท' if names else 'ยังไม่ได้เลือกชีท'),
                            fg=C['fg'] if names else C['dim'])
        if self._pending is not None:
            try:
                self.top.after_cancel(self._pending)
            except Exception:
                pass
        self._pending = self.top.after(300, self._request)

    def _request(self):
        self._pending = None
        key = self._key()
        if not key:
            self._clear()
            self.info.config(text='เลือกชีทอย่างน้อยหนึ่งชีท')
            return
        if key in self.cache:
            rows, note = self.cache[key]
            self._show(rows, note)
            return
        self._clear()
        names = self._selected()
        self.info.config(text=(f'กำลังอ่าน {len(names)} ชีท…' if len(names) > 1
                               else f'กำลังอ่านชีท “{names[0]}” …'))
        if key in self.inflight:
            return
        self.inflight.add(key)
        threading.Thread(target=self._rows_worker, args=(key, names), daemon=True).start()

    def _rows_worker(self, key, names):
        rows, note = [], ''
        try:
            with self.lock:
                seen = set()          # ตัด ItemKind ซ้ำข้ามชีทให้เหลือตัวเดียว
                hit = []
                for k, name in enumerate(names, 1):
                    if len(names) > 1:
                        self.q.put(('info', 'กำลังอ่าน %d/%d ชีท…  %s  (ได้แล้ว %d รายการ)'
                                    % (k, len(names), name, len(rows))))
                    got = parse_master_rows(read_sheet_rows(self.wb, name), seen)
                    if got:
                        hit.append('%s (%d)' % (name, len(got)))
                    rows.extend(got)
                if len(names) == 1:
                    note = 'ชีท “%s”' % names[0]
                else:
                    note = 'รวม %d ชีท: %s' % (len(names), ', '.join(hit) or '-')
        except Exception as ex:
            self.q.put(('info', 'อ่านไม่ได้: ' + str(ex)))
        self.q.put(('rows', key, rows, note))

    # ---------- แสดงผล ----------
    def _clear(self):
        self.rows = []
        self.tree.delete(*self.tree.get_children())
        self.count_lbl.config(text='')
        self.ok_btn.config(state='disabled', bg=C['input'], fg=C['dim'])

    def _why(self):
        """คลิกแถวเหลือง แล้วบอกว่าเตือนเพราะอะไร"""
        s = self.tree.selection()
        if not s:
            return
        i = self.tree.index(s[0])
        if 0 <= i < len(self.rows):
            w = self.rows[i].get('warn') or []
            if w:
                self.warn_lbl.config(text='⚠  ' + ' · '.join(w), fg=C['warn'])
            else:
                self.warn_lbl.config(text='', fg=C['dim'])

    def _show(self, rows, note):
        self.rows = rows
        nwarn = mark_suspects(rows)
        self.tree.delete(*self.tree.get_children())
        for r in rows[:400]:
            dur = 'ถาวร' if r['dur'] == '' else (r['dur'] + ' วัน')
            trade = {'yes': 'ได้', 'no': 'ไม่ได้', 'any': '— ไม่กรอง —'}.get(r['trade'], r['trade'])
            self.tree.insert('', 'end', tags=('warn',) if r.get('warn') else (),
                             values=('⚠' if r.get('warn') else '', r['kind'], r['disp'],
                                     dur, trade, r['qty'] or '—'))
        has_move = any(r.get('has_move') for r in rows)
        extra = '' if has_move else '   (ชีทนี้ไม่มีคอลัมน์ Itemmove → ไม่กรองแลกเปลี่ยน)'
        self.info.config(text=f'{note} · พบ {len(rows)} รายการ{extra}')
        if nwarn:
            self.warn_lbl.config(
                text=f'⚠  มี {nwarn} แถวที่หน้าตาไม่เหมือนไอเทมทั่วไป (แถวสีเหลือง) — '
                     'ยังนำเข้าให้ครบ คลิกที่แถวเพื่อดูว่าเตือนเพราะอะไร',
                fg=C['warn'])
        else:
            self.warn_lbl.config(text='', fg=C['dim'])
        self.count_lbl.config(text=f'จะนำเข้า {len(rows)} รายการ'
                              + (f'  (น่าสงสัย {nwarn})' if nwarn else '')
                              + ('   (แสดงตัวอย่าง 400 แถวแรก)' if len(rows) > 400 else ''))
        self.ok_btn.config(state='normal' if rows else 'disabled',
                           bg=C['accent'] if rows else C['input'],
                           fg='white' if rows else C['dim'])

    def _close_wb(self):
        try:
            if self.wb:
                self.wb.close()
        except Exception:
            pass
        self.wb = None

    def _ok(self):
        self.result = list(self.rows)
        self._close_wb()
        self.top.destroy()

    def _cancel(self):
        self.result = None
        self._close_wb()
        self.top.destroy()


# ============================================================================
#  [5.6] ระบบตรวจสอบตัวเอง (Self Test)
#        แบ่งเป็น 2 ชุด
#          A. ตรวจตรรกะ  — ไม่ต้องเปิดเว็บ เช็กว่ากฎการอ่านชีทยังถูกต้อง
#          B. ตรวจเว็บ    — เปิดเว็บจริง เช็กว่าหน้าตาเว็บยังตรงกับ CFG ไหม
#        ถ้าเว็บเปลี่ยน จะได้รู้ทันทีว่า "พังตรงไหน" และ "ต้องแก้ค่าตัวไหน"
# ============================================================================
def _t(ok, name, detail='', fix=''):
    return {'ok': bool(ok), 'name': name, 'detail': str(detail), 'fix': fix}


def run_logic_tests():
    """ตรวจกฎการอ่านข้อมูลทั้งหมด โดยไม่ต้องต่อเน็ต"""
    r = []

    # ---- กติกาช่องระยะเวลา ----
    r.append(_t(dur_from_text('') == '', 'ระยะเวลา: เว้นว่าง = ถาวร', repr(dur_from_text(''))))
    r.append(_t(dur_from_text('ถาวร') == '', 'ระยะเวลา: "ถาวร" = ถาวร', repr(dur_from_text('ถาวร'))))
    r.append(_t(dur_from_text('any') == 'any', 'ระยะเวลา: "any" = ไม่กรอง', repr(dur_from_text('any'))))
    r.append(_t(dur_from_text('15 วัน') == '15', 'ระยะเวลา: "15 วัน" = 15', repr(dur_from_text('15 วัน'))))

    # ---- ตัวเลขจาก Excel ----
    r.append(_t(num_str(112247.0) == '112247', 'ตัดทศนิยม .0 ที่ Excel ติดมา', num_str(112247.0)))

    # ---- Yes/No ----
    r.append(_t(all(norm_bool(v) == 'yes' for v in ('Yes', 'y', '1', 'TRUE')),
                'อ่านค่า Yes ได้ทุกแบบ'))
    r.append(_t(all(norm_bool(v) == 'no' for v in ('No', 'n', '0', 'false')),
                'อ่านค่า No ได้ทุกแบบ'))
    r.append(_t(norm_bool('') == 'any' and norm_bool('มั่ว') == 'any',
                'ค่าว่าง/อ่านไม่ออก = ไม่กรอง'))

    # ---- อ่านชีทต้นฉบับ (กฎเดียวกับเครื่องมือเดิม) ----
    rows = [
        ['400', 'fdItemNum', 'fdPosition', 'fdItemKind', 'Rank',
         'Display Name', 'ระยะเวลาของขวัญ', 'Amt', 'ราคา'],
        ['', '112247', '9999', '2648', 'SS', 'New.1 ตั๋วเปลี่ยนตำนาน II 1 ชิ้น', 'ถาวร', '1', '150'],
        ['', '128172', '250', '5485', 'S', 'กล่องอิลลิเฟียร์ 5 ชิ้น', 'ถาวร', '1', '95'],
        ['', '999001', '250', '5486', 'A', 'หมวกทดสอบ 3 ชิ้น', '7 วัน', '1', '10'],
        ['', '', '', '', '', '', '', '', '460'],
        ['', 'ขอรางวัลเริ่ดๆ เหมือนเดิมค่า'],
    ]
    got = parse_master_rows(rows)
    r.append(_t(len(got) == 3, 'อ่านชีทตัวอย่างได้ 3 รายการ (ตัดแถวยอดรวม/หมายเหตุ)', len(got)))
    if len(got) == 3:
        a, b, c = got
        r.append(_t(a['kind'] == '112247', 'ใช้คอลัมน์ fdItemNum เป็น ItemKind', a['kind']))
        r.append(_t(a['qty'] == '1' and b['qty'] == '5' and c['qty'] == '3',
                    'ดึงจำนวนจากเลขหน้าคำว่า "ชิ้น"',
                    f"{a['qty']}/{b['qty']}/{c['qty']}"))
        r.append(_t(b['disp'] == 'กล่องอิลลิเฟียร์', 'ตัดคำว่า "X ชิ้น" ออกจากชื่อ', b['disp']))
        r.append(_t(a['dur'] == '' and c['dur'] == '7',
                    'อ่านระยะเวลา: ถาวร / 7 วัน', f"{a['dur']!r}/{c['dur']!r}"))
        r.append(_t(a['trade'] == 'any', 'ไม่มีคอลัมน์ Itemmove = ไม่กรองแลกเปลี่ยน', a['trade']))

    # ---- มีคอลัมน์ Itemmove ----
    rows2 = [
        ['fdItemNum', 'Display Name', 'ระยะเวลาของขวัญ', 'Itemmove'],
        ['111111', 'ของทดสอบ 2 ชิ้น', 'ถาวร', 'Yes'],
        ['222222', 'ของทดสอบสอง 1 ชิ้น', 'ถาวร', ''],
    ]
    g2 = parse_master_rows(rows2)
    r.append(_t(len(g2) == 2 and g2[0]['trade'] == 'yes' and g2[1]['trade'] == 'no',
                'มีคอลัมน์ Itemmove: Yes = ได้ / ว่าง = ไม่ได้',
                [x['trade'] for x in g2]))

    # ---- ตัด ID ซ้ำ ----
    seen = set()
    d1 = parse_master_rows(rows, seen)
    d2 = parse_master_rows(rows, seen)
    r.append(_t(len(d1) == 3 and len(d2) == 0, 'ตัด ID ซ้ำข้ามชีทได้', f'{len(d1)} / {len(d2)}'))

    # ---- เงื่อนไข deep check ----
    item7 = {'duration': '7', 'trade': False, 'qty': '10'}
    r.append(_t(match_deep(item7, {'dur': '7', 'trade': 'any', 'qty': ''})[0],
                'deep check: ระยะเวลาตรง = ผ่าน'))
    r.append(_t(not match_deep(item7, {'dur': '15', 'trade': 'any', 'qty': ''})[0],
                'deep check: ระยะเวลาไม่ตรง = ไม่ผ่าน'))
    r.append(_t(match_deep({'duration': '0', 'trade': None, 'qty': ''},
                           {'dur': '', 'trade': 'any', 'qty': ''})[0],
                'deep check: ระยะเวลา 0 บนเว็บ = ถาวร'))
    r.append(_t(not match_deep(item7, {'dur': 'any', 'trade': 'yes', 'qty': ''})[0],
                'deep check: แลกเปลี่ยนไม่ตรง = ไม่ผ่าน'))
    r.append(_t(not match_deep(item7, {'dur': 'any', 'trade': 'any', 'qty': '5'})[0],
                'deep check: จำนวนไม่ตรง = ไม่ผ่าน'))

    # ---- ทางลัดใช้ฟิลเตอร์หน้า list ----
    r.append(_t(bool(duration_only_numeric({'dur': '15', 'trade': 'any', 'qty': ''})),
                'ใช้ทางลัดฟิลเตอร์เมื่อกรองแค่ระยะเวลา'))
    r.append(_t(not duration_only_numeric({'dur': '15', 'trade': 'yes', 'qty': ''}),
                'ไม่ใช้ทางลัดเมื่อมีเงื่อนไขอื่นด้วย'))

    # ---- ระบบบันทึกการใช้งาน ----
    try:
        before = len(read_events())
        log_event('selftest_ping')
        after = len(read_events())
        ok_log = after == before + 1
        detail = f'{after} เหตุการณ์ในไฟล์'
    except Exception as ex:
        ok_log, detail = False, str(ex)[:80]
    r.append(_t(ok_log, 'บันทึกการใช้งานลงไฟล์ได้', f'{detail}  ({LOCAL_LOG_DIR})',
                'เช็กสิทธิ์เขียนโฟลเดอร์ที่ตั้งไว้'))
    st = read_stats()
    r.append(_t(st.get('total_events', 0) > 0, 'ตัวนับสถิติทำงาน',
                f"รวม {st.get('total_events', 0)} ครั้ง · เริ่มเก็บ {st.get('first_used', '-')}"))
    if CENTRAL_LOG_DIR:
        r.append(_t(os.path.isdir(CENTRAL_LOG_DIR), 'โฟลเดอร์กลางของทีมเขียนได้',
                    CENTRAL_LOG_DIR, 'โฟลเดอร์หายไป — ตั้งใหม่ที่ปุ่มข้างบน'))
    else:
        r.append(_t(True, 'โฟลเดอร์กลางของทีม',
                    'ยังไม่ได้ตั้ง — เก็บเฉพาะในเครื่องนี้ (ตั้งได้ที่ปุ่มข้างบน)'))

    # ---- กติกาหน้า "สร้างไอเทม" ----
    a = split_name_qty('กล่องแอนิมอล 1 ชิ้น')
    r.append(_t(a == ('กล่องแอนิมอล', '1', True),
                'สร้างไอเทม: ตัด "จำนวนชิ้น" ออกจากชื่อ', repr(a),
                'กฎเดิม: จำนวนชิ้นไม่ติดไปในชื่อ'))
    b = split_name_qty('แพ็กเบ็ดตกปลา 2 ชั่วโมง 3 ชิ้น')
    r.append(_t(b == ('แพ็กเบ็ดตกปลา 2 ชั่วโมง', '3', True),
                'สร้างไอเทม: เก็บ "ชม." ไว้ในชื่อ ตัดแต่จำนวน', repr(b),
                'กฎเดิม: ชม. อยู่ในชื่อได้ จำนวนอยู่คนละช่อง'))
    c = split_name_qty('ค้อนซ่อมแซม')
    r.append(_t(c == ('ค้อนซ่อมแซม', '1', False),
                'สร้างไอเทม: ชื่อที่ไม่มี "ชิ้น" คงชื่อเดิม จำนวน 1', repr(c)))
    d1 = make_item_name('กล่องแอนิมอล', '(Gift)', '3')
    r.append(_t(d1 == 'กล่องแอนิมอล (Gift) (3 วัน)',
                'สร้างไอเทม: ประกอบชื่อ + คำต่อท้าย + (X วัน)', d1))
    e1 = row_to_item({'kind': '1', 'cname': 'ค้อน', 'amount': '1',
                      'has_qty': False, 'dur': '5', 'trade': 'yes'}, 'GENERAL', '(Gift)')
    r.append(_t(e1['name'] == 'ค้อน',
                'สร้างไอเทม: ชื่อไม่มี "ชิ้น" -> ไม่เติมคำต่อท้าย/วัน', e1['name'],
                'กฎเดิมเติมเฉพาะชื่อที่มี "ชิ้น"'))

    # ---- ระบบเตือนของแปลกปลอม (เตือนอย่างเดียว ห้ามตัดแถวทิ้ง) ----
    junk = [{'kind': '128107', 'raw': 'ก'}, {'kind': '128108', 'raw': 'ข'},
            {'kind': '128109', 'raw': 'ค'}, {'kind': '2', 'raw': 'แก้ UID เรียบร้อย'}]
    nw = mark_suspects(junk)
    r.append(_t(nw == 1 and len(junk) == 4,
                'เตือนแถวแปลกปลอมได้ โดยไม่ตัดแถวทิ้ง',
                'เตือน %d แถว · เหลือ %d แถว' % (nw, len(junk))))
    longest = 'แพ็กเบ็ดตกปลาเทลส์รันเนอร์ 2 ชั่วโมง (Gift) (30 วัน)'
    r.append(_t(suspect_reasons({'kind': '128909', 'raw': longest}, 6) == [],
                'ชื่อไอเทมจริงที่ยาวที่สุด ไม่ถูกเตือนผิด',
                '%d ตัวอักษร (เพดาน %d)' % (len(longest), MAX_SANE_NAME)))
    junk[3]['kind'] = '128110'
    r.append(_t(mark_suspects(junk) == 0,
                'แก้ ItemKind แล้วคำเตือนหายเอง'))

    # ---- สภาพแวดล้อม ----
    r.append(_t(XLSX_OK, 'อ่านไฟล์ Excel ได้ (openpyxl)', 'ok' if XLSX_OK else 'ไม่มี openpyxl'))
    chrome = find_chrome_exe()
    r.append(_t(chrome, 'เจอ Google Chrome ในเครื่อง', chrome or 'ไม่พบ',
                'ติดตั้ง Chrome ก่อนใช้งาน'))
    return r


JS_HEADERS = """() => [...document.querySelectorAll('table thead th')]
                        .map(x => (x.innerText || '').trim())"""


async def run_web_tests(page, log=None):
    """เปิดเว็บจริงแล้วตรวจว่าหน้าตายังตรงกับ CFG ไหม"""
    r = []

    def say(msg):
        if log:
            log(msg)

    # --- W1 เปิดหน้า Item ---
    try:
        await page.goto(ITEM_LIST_URL, wait_until='domcontentloaded', timeout=45000)
        await page.wait_for_timeout(1800)
        r.append(_t(True, 'เปิดหน้า Shop > Item ได้', page.url.split('?')[0]))
    except Exception as ex:
        r.append(_t(False, 'เปิดหน้า Shop > Item ได้', str(ex)[:90], 'เช็กอินเทอร์เน็ต / ITEM_LIST_URL'))
        return r
    say('  เปิดหน้า Item แล้ว')

    # --- W2 ล็อกอินอยู่ไหม ---
    has_search = await page.locator(SEL['search_input']).count() > 0
    r.append(_t(has_search, 'ล็อกอินอยู่ และเจอช่องค้นหา',
                'เจอ' if has_search else 'ไม่เจอช่องค้นหา',
                'กดปุ่ม "เปิดหน้า Login" แล้วล็อกอินก่อน'))
    if not has_search:
        return r

    # --- W3 หัวตาราง ---
    try:
        heads = await page.evaluate(JS_HEADERS)
    except Exception:
        heads = []
    ok_head = len(heads) >= 4 and 'Aztek Item Id' in ' '.join(heads) and 'ItemKind' in ' '.join(heads)
    r.append(_t(ok_head, 'ตารางมีคอลัมน์ Aztek Item Id และ ItemKind',
                ' | '.join(heads) or 'อ่านหัวตารางไม่ได้', 'แก้ COL ใน CONFIG'))

    # --- W4 อ่านแถวได้ ---
    rows = await page.evaluate(JS_READ_ROWS, COL)
    r.append(_t(len(rows) > 0, 'อ่านแถวในตารางได้', f'{len(rows)} แถว', 'แก้ SEL["row"] หรือ COL'))
    if not rows:
        return r
    first = rows[0]
    r.append(_t(re.fullmatch(r'\d+', first['id'] or '') and first['name'],
                'คอลัมน์ ID เป็นตัวเลข และมีชื่อไอเทม',
                f"id={first['id']} name={first['name'][:24]}", 'ลำดับคอลัมน์ COL เปลี่ยน'))

    # --- W5 ปุ่มค้นหา + ฟิลเตอร์ ---
    r.append(_t(await page.locator(SEL['search_button']).count() > 0,
                'เจอปุ่มค้นหา', SEL['search_button'], 'แก้ SEL["search_button"]'))
    r.append(_t(await page.locator(SEL['filter_duration']).count() > 0,
                'เจอฟิลเตอร์ DurationIndex', SEL['filter_duration'],
                'แก้ SEL["filter_duration"]'))
    nxt = await page.locator(f'button:has-text("{SEL["next_text"]}")').count()
    r.append(_t(nxt > 0, 'เจอปุ่มเปลี่ยนหน้า (Next)', f'{nxt} ปุ่ม', 'แก้ SEL["next_text"]'))

    # --- W6 ค้นหาจริงแล้วต้องเจอ ---
    kind = first['kind'] or first['id']
    say(f'  ทดลองค้นหา {kind}')
    try:
        await page.fill(SEL['search_input'], str(kind))
        await page.locator(SEL['search_button']).first.click()
        await page.wait_for_timeout(1600)
        found = await page.evaluate(JS_READ_ROWS, COL)
        hit = any(x['kind'] == first['kind'] or x['id'] == first['id'] for x in found)
        r.append(_t(hit, f'ค้นหา {kind} แล้วเจอไอเทมที่ต้องการ', f'ได้ {len(found)} แถว',
                    'ช่องค้นหาหรือปุ่มค้นหาเปลี่ยนไป'))
    except Exception as ex:
        r.append(_t(False, f'ค้นหา {kind} แล้วเจอไอเทมที่ต้องการ', str(ex)[:90]))
        return r

    # --- W7 เข้าหน้ารายละเอียด ---
    target = first['id']
    try:
        clicked = await page.evaluate(JS_CLICK_ROW, str(target))
        if clicked:
            await page.wait_for_selector(SEL['d_name'], timeout=15000)
            await page.wait_for_timeout(500)
        r.append(_t(clicked, 'คลิกเข้าหน้ารายละเอียดไอเทมได้',
                    page.url.split('/')[-1], 'ลิงก์ในตารางเปลี่ยนรูปแบบ'))
        if not clicked:
            return r
    except Exception as ex:
        r.append(_t(False, 'คลิกเข้าหน้ารายละเอียดไอเทมได้', str(ex)[:90]))
        return r

    # --- W8 ฟิลด์ในหน้ารายละเอียด ---
    for key, label in (('d_name', 'ชื่อไอเท็ม'), ('d_kind', 'game_item_id (ItemKind)'),
                       ('d_price', 'Price'), ('d_duration', 'ระยะเวลาไอเท็ม (วัน)'),
                       ('d_qty', 'จำนวน')):
        n = await page.locator(SEL[key]).count()
        r.append(_t(n > 0, f'เจอช่อง {label}', SEL[key], f'แก้ SEL["{key}"]'))

    detail = await page.evaluate(JS_READ_DETAIL, {
        'name': SEL['d_name'], 'kind': SEL['d_kind'], 'price': SEL['d_price'],
        'duration': SEL['d_duration'], 'qty': SEL['d_qty'],
        'tradeLabel': SEL['d_trade_label'],
    })
    r.append(_t(detail.get('trade') is not None, 'อ่านค่า "แลกเปลี่ยนได้" ได้',
                detail.get('trade'), 'แก้ SEL["d_trade_label"]'))

    # --- W9 ข้อมูลตรงกันระหว่างตารางกับหน้ารายละเอียด (ตรวจ mapping ทั้งเส้น) ---
    same = str(detail.get('kind') or '').strip() == str(first['kind'] or '').strip()
    r.append(_t(same, 'ItemKind ในตาราง ตรงกับในหน้ารายละเอียด',
                f"ตาราง={first['kind']}  รายละเอียด={detail.get('kind')}",
                'คอลัมน์ ItemKind ในตารางอาจสลับตำแหน่ง — แก้ COL'))
    return r

# ============================================================================
#  [6] หน้าต่างโปรแกรม
# ============================================================================
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(f'TR Plus Ultra  —  v{APP_VERSION}')
        self.root.configure(bg=C['bg'])
        self.root.geometry('1080x720')
        self.prefs = load_prefs()

        self.results = []
        self.not_found = []
        self.imported = []
        self.imp_src = ''
        self.running = False
        self.cancel = False

        load_local_log_dir()
        load_central_dir()
        self._build_ui()
        trim_history()
        log_event('app_open', launcher=globals().get('TRPU_LAUNCHER', ''))
        for t in pending_items():
            self.log('⚠  ' + t, 'WARN')

    # ---------- UI ----------
    def _build_ui(self):
        top = tk.Frame(self.root, bg=C['card'], height=54)
        top.pack(fill='x')
        top.pack_propagate(False)
        tk.Label(top, text='TR Plus Ultra', bg=C['card'], fg=C['fg'],
                 font=('Segoe UI', 14, 'bold')).pack(side='left', padx=16)
        tk.Label(top, text=f'v{APP_VERSION}', bg=C['card'], fg=C['dim'],
                 font=('Segoe UI', 9)).pack(side='left')
        tk.Button(top, text='🔓  เปิดหน้า Login', bg=C['input'], fg=C['fg'], bd=0,
                  font=FM, cursor='hand2', activebackground=C['line'],
                  command=self.open_login).pack(side='right', padx=16, ipadx=12, ipady=5)

        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass
        style.configure('TNotebook', background=C['bg'], borderwidth=0)
        style.configure('TNotebook.Tab', background=C['card'], foreground=C['dim'],
                        padding=(18, 8), font=FM, borderwidth=0)
        style.map('TNotebook.Tab', background=[('selected', C['bg'])],
                  foreground=[('selected', C['fg'])])

        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill='both', expand=True, padx=10, pady=(8, 10))

        self.tab_search = tk.Frame(self.nb, bg=C['bg'])
        self.tab_create = tk.Frame(self.nb, bg=C['bg'])
        self.tab_result = tk.Frame(self.nb, bg=C['bg'])
        self.tab_log = tk.Frame(self.nb, bg=C['bg'])
        self.tab_check = tk.Frame(self.nb, bg=C['bg'])
        self.nb.add(self.tab_search, text='🔍  ค้นหา')
        self.nb.add(self.tab_create, text='➕  สร้าง Item')
        self.nb.add(self.tab_result, text='📋  ผลลัพธ์')
        self.nb.add(self.tab_log, text='📜  Log')
        self.nb.add(self.tab_check, text='🩺  ตรวจระบบ')

        self._build_search()
        self._build_create()
        self._build_result()
        self._build_log()
        self._build_check()

    def _card(self, parent, title):
        outer = tk.LabelFrame(parent, text='  ' + title + '  ', bg=C['bg'], fg=C['dim'],
                              font=('Segoe UI', 9, 'bold'), bd=1,
                              relief='solid', highlightbackground=C['line'])
        outer.pack(fill='x', padx=14, pady=(12, 0))
        inner = tk.Frame(outer, bg=C['bg'])
        inner.pack(fill='x', padx=12, pady=10)
        return inner

    def _entry(self, parent, width=22):
        e = tk.Entry(parent, bg=C['input'], fg=C['fg'], insertbackground=C['fg'],
                     bd=0, font=FM, width=width, relief='flat',
                     highlightthickness=1, highlightbackground=C['line'],
                     highlightcolor=C['accent'])
        return e

    def _btn(self, parent, text, cmd, primary=False, width=None):
        b = tk.Button(parent, text=text, command=cmd, bd=0, cursor='hand2', font=FB if primary else FM,
                      bg=C['accent'] if primary else C['input'],
                      fg='white' if primary else C['fg'],
                      activebackground='#3d55cf' if primary else C['line'],
                      activeforeground='white')
        if width:
            b.config(width=width)
        return b

    def _build_search(self):
        p = self.tab_search

        s1 = self._card(p, 'ค้นหาทีละตัว')
        tk.Label(s1, text='ItemKind / Aztek Item Id', bg=C['bg'], fg=C['dim'],
                 font=('Segoe UI', 9)).grid(row=0, column=0, sticky='w')
        tk.Label(s1, text='ชื่อไอเทม (ไม่บังคับ)', bg=C['bg'], fg=C['dim'],
                 font=('Segoe UI', 9)).grid(row=0, column=1, sticky='w', padx=(12, 0))
        self.v_kind = self._entry(s1, 24)
        self.v_kind.grid(row=1, column=0, sticky='w', ipady=4)
        self.v_name = self._entry(s1, 34)
        self.v_name.grid(row=1, column=1, sticky='w', padx=(12, 0), ipady=4)
        self.btn_single = self._btn(s1, '🔍  ค้นหา', self.run_single, primary=True)
        self.btn_single.grid(row=1, column=2, padx=(14, 0), ipadx=18, ipady=3)

        s2 = self._card(p, 'ค้นหาหลายตัวจาก Excel')
        self._btn(s2, '⬇  โหลด Template', self.download_template).grid(row=0, column=0, ipadx=10, ipady=4)
        self._btn(s2, '📂  เลือกไฟล์', self.pick_file).grid(row=0, column=1, padx=8, ipadx=10, ipady=4)
        self.lbl_file = tk.Label(s2, text='ยังไม่ได้เลือกไฟล์', bg=C['bg'], fg=C['dim'], font=('Segoe UI', 9))
        self.lbl_file.grid(row=0, column=2, padx=12, sticky='w')
        self.btn_multi = self._btn(s2, '🔍  ค้นหาทั้งหมด', self.run_multi, primary=True)
        self.btn_multi.grid(row=0, column=3, padx=(8, 0), ipadx=14, ipady=3)
        self.btn_multi.config(state='disabled')

        # ---- ปุ่มกาง/พับ รายการที่จะค้นหา ----
        self.imp_open = False
        self.btn_imp = self._btn(s2, '▸  ดูรายการที่จะค้นหา', self.toggle_imp)
        self.btn_imp.grid(row=1, column=0, columnspan=2, sticky='w', pady=(10, 0),
                          ipadx=8, ipady=3)
        self.btn_imp.config(state='disabled')
        self.lbl_imp = tk.Label(s2, text='', bg=C['bg'], fg=C['warn'], font=('Segoe UI', 9),
                                anchor='w')
        self.lbl_imp.grid(row=1, column=2, columnspan=2, sticky='w', padx=12, pady=(10, 0))

        # ---- แผงรายการ (ซ่อนไว้ก่อน กางเมื่อกดปุ่ม) ----
        self.imp_box = tk.LabelFrame(p, text='  รายการที่จะค้นหา  ', bg=C['bg'], fg=C['dim'],
                                     font=('Segoe UI', 9, 'bold'), bd=1, relief='solid')
        ib = tk.Frame(self.imp_box, bg=C['bg'])
        ib.pack(fill='both', expand=True, padx=12, pady=10)
        ibar = tk.Frame(ib, bg=C['bg'])
        ibar.pack(fill='x')
        self._btn(ibar, '🗑  เอาออกจากรายการ', self.imp_del).pack(side='left', ipadx=8, ipady=3)
        self._btn(ibar, '📋  คัดลอกเลขทั้งหมด', self.imp_copy).pack(side='left', padx=8,
                                                                    ipadx=8, ipady=3)
        self.imp_cnt = tk.Label(ibar, text='', bg=C['bg'], fg=C['dim'], font=FM)
        self.imp_cnt.pack(side='right')

        itw = tk.Frame(ib, bg=C['bg'])
        itw.pack(fill='both', expand=True, pady=(8, 0))
        icols = ('n', 'w', 'kind', 'name', 'dur', 'trade', 'qty')
        self.imp_tree = ttk.Treeview(itw, columns=icols, show='headings',
                                     style='TR.Treeview', height=6, selectmode='extended')
        for c, t, w in (('n', '#', 40), ('w', '', 26), ('kind', 'ItemKind', 95),
                        ('name', 'ชื่อไอเทม', 360), ('dur', 'ระยะเวลา', 90),
                        ('trade', 'แลกเปลี่ยน', 95), ('qty', 'จำนวน', 70)):
            self.imp_tree.heading(c, text=t)
            self.imp_tree.column(c, width=w, anchor='w')
        self.imp_tree.tag_configure('warn', background='#3a3018', foreground='#e3b341')
        isb = ttk.Scrollbar(itw, orient='vertical', command=self.imp_tree.yview)
        self.imp_tree.configure(yscrollcommand=isb.set)
        self.imp_tree.pack(side='left', fill='both', expand=True)
        isb.pack(side='right', fill='y')
        self.imp_why = tk.Label(ib, text='', bg=C['bg'], fg=C['warn'], font=('Segoe UI', 9),
                                anchor='w', justify='left', wraplength=940)
        self.imp_why.pack(fill='x', pady=(6, 0))
        self.imp_tree.bind('<<TreeviewSelect>>', lambda e: self._imp_why())

        s3 = self._card(p, 'Deep Check  (เข้าไปดูค่าจริงในหน้ารายละเอียด)')
        self.v_deep = tk.BooleanVar(value=bool(self.prefs.get('deep', False)))
        tk.Checkbutton(s3, text='เปิด Deep Check', variable=self.v_deep, bg=C['bg'], fg=C['fg'],
                       selectcolor=C['input'], activebackground=C['bg'], activeforeground=C['fg'],
                       font=FM, bd=0, highlightthickness=0).grid(row=0, column=0, sticky='w')

        tk.Label(s3, text='ระยะเวลา (วัน)', bg=C['bg'], fg=C['dim'],
                 font=('Segoe UI', 9)).grid(row=1, column=0, sticky='w', pady=(8, 0))
        tk.Label(s3, text='แลกเปลี่ยนได้', bg=C['bg'], fg=C['dim'],
                 font=('Segoe UI', 9)).grid(row=1, column=1, sticky='w', padx=(12, 0), pady=(8, 0))
        tk.Label(s3, text='จำนวน (ต้องตรงเป๊ะ)', bg=C['bg'], fg=C['dim'],
                 font=('Segoe UI', 9)).grid(row=1, column=2, sticky='w', padx=(12, 0), pady=(8, 0))

        self.v_dur = self._entry(s3, 18)
        self.v_dur.insert(0, self.prefs.get('dur', 'any'))
        self.v_dur.grid(row=2, column=0, sticky='w', ipady=4)
        self.v_trade = ttk.Combobox(s3, values=['any', 'yes', 'no'], width=10, state='readonly', font=FM)
        self.v_trade.set(self.prefs.get('trade', 'any'))
        self.v_trade.grid(row=2, column=1, sticky='w', padx=(12, 0))
        self.v_qty = self._entry(s3, 18)
        self.v_qty.insert(0, self.prefs.get('qty', ''))
        self.v_qty.grid(row=2, column=2, sticky='w', padx=(12, 0), ipady=4)

        tk.Label(s3, text='ระยะเวลา: any = ไม่กรอง · เว้นว่าง = เฉพาะไอเทมถาวร (ระยะเวลา = 0) · ตัวเลข = จำนวนวันนั้น\n'
                          'ถ้ากรองแค่ระยะเวลาเป็นตัวเลข โปรแกรมจะใช้ฟิลเตอร์บนหน้า list แทน เร็วกว่ามาก',
                 bg=C['bg'], fg=C['dim'], font=('Segoe UI', 8), justify='left'
                 ).grid(row=3, column=0, columnspan=4, sticky='w', pady=(8, 0))

        s4 = tk.Frame(p, bg=C['bg'])
        self.s_tail = s4
        s4.pack(fill='x', padx=14, pady=(14, 0))
        self.v_headless = tk.BooleanVar(value=bool(self.prefs.get('headless', False)))
        tk.Checkbutton(s4, text='ซ่อนหน้าต่าง Chrome ตอนทำงาน', variable=self.v_headless,
                       bg=C['bg'], fg=C['dim'], selectcolor=C['input'], activebackground=C['bg'],
                       activeforeground=C['fg'], font=('Segoe UI', 9), bd=0,
                       highlightthickness=0).pack(side='left')
        self.btn_cancel = self._btn(s4, '■  ยกเลิก', self.do_cancel)
        self.btn_cancel.config(state='disabled', fg=C['err'])
        self.btn_cancel.pack(side='right', ipadx=14, ipady=4)

        s5 = tk.Frame(p, bg=C['bg'])
        s5.pack(fill='x', padx=14, pady=(12, 0))
        self.progress = ttk.Progressbar(s5, mode='determinate', maximum=100)
        self.progress.pack(fill='x')
        self.lbl_stat = tk.Label(s5, text='', bg=C['bg'], fg=C['dim'], font=('Segoe UI', 9), anchor='w')
        self.lbl_stat.pack(fill='x', pady=(5, 0))

    # ========================================================================
    #  แท็บ "สร้าง Item"
    #  แนวคิด: โยนไฟล์ต้นฉบับเข้าไป -> ได้คิวไอเทม -> แก้ได้ทุกช่อง -> ยิงเข้าเว็บ
    #  ไม่ต้องกรอก template เองเหมือนแต่ก่อน
    # ========================================================================
    def _build_create(self):
        p = self.tab_create
        self.cq = []                       # คิวไอเทมที่จะสร้าง
        self.c_running = False
        self.c_cancel = False

        # ---------- ค่าเริ่มต้นที่ใช้กับทุกแถว ----------
        s1 = self._card(p, 'ค่าเริ่มต้น (ใช้ตอนนำเข้าไฟล์ — แก้รายตัวทีหลังได้)')
        self.cv_type = tk.StringVar(value=self.prefs.get('c_type', DEFAULT_TYPE))
        self.cv_suffix = tk.StringVar(value=self.prefs.get('c_suffix', DEFAULT_SUFFIX))
        self.cv_price = tk.StringVar(value=self.prefs.get('c_price', ''))
        self.cv_mail = tk.StringVar(value=self.prefs.get('c_mail', ''))

        for col, (lbl, var, w) in enumerate((
                ('ประเภทไอเทม', self.cv_type, 14),
                ('คำต่อท้ายชื่อ', self.cv_suffix, 12),
                ('Price', self.cv_price, 10),
                ('หัวข้อจดหมาย', self.cv_mail, 20))):
            tk.Label(s1, text=lbl, bg=C['bg'], fg=C['dim'],
                     font=('Segoe UI', 9)).grid(row=0, column=col * 2, sticky='w', padx=(0 if col == 0 else 14, 6))
            e = self._entry(s1, width=w)
            e.config(textvariable=var)
            e.grid(row=0, column=col * 2 + 1, sticky='w', ipady=3)

        tk.Label(s1, text='คำต่อท้ายกับ “(X วัน)” จะเติมให้เฉพาะชื่อที่มีคำว่า “ชิ้น” เท่านั้น — '
                          'จำนวนชิ้นจะถูก “ตัดออกจากชื่อ” แต่ ชม. ยังอยู่ในชื่อเหมือนเดิม',
                 bg=C['bg'], fg=C['dim'], font=('Segoe UI', 8), wraplength=940,
                 justify='left').grid(row=1, column=0, columnspan=8, sticky='w', pady=(8, 0))

        # ---------- คิว ----------
        s2 = self._card(p, 'คิวไอเทมที่จะสร้าง')
        bar = tk.Frame(s2, bg=C['bg'])
        bar.pack(fill='x')
        self._btn(bar, '📂  นำเข้าไฟล์ต้นฉบับ (.xlsx)', self.c_import,
                  primary=True).pack(side='left', ipadx=10, ipady=4)
        self._btn(bar, '➕  เพิ่มเอง', self.c_add).pack(side='left', padx=(8, 0), ipadx=8, ipady=4)
        self._btn(bar, '✏  แก้ไข', self.c_edit).pack(side='left', padx=(8, 0), ipadx=8, ipady=4)
        self._btn(bar, '🗑  ลบที่เลือก', self.c_del).pack(side='left', padx=(8, 0), ipadx=8, ipady=4)
        self._btn(bar, 'ล้างคิว', self.c_clear).pack(side='left', padx=(8, 0), ipadx=8, ipady=4)
        self.c_count = tk.Label(bar, text='คิวว่าง', bg=C['bg'], fg=C['dim'], font=FM)
        self.c_count.pack(side='right')

        tw = tk.Frame(s2, bg=C['bg'])
        tw.pack(fill='both', expand=True, pady=(10, 0))
        cols = ('n', 'w', 'kind', 'name', 'type', 'price', 'dur', 'qty', 'trade')
        self.c_tree = ttk.Treeview(tw, columns=cols, show='headings',
                                   style='TR.Treeview', height=8)
        for c, t, w in (('n', '#', 38), ('w', '', 26), ('kind', 'ItemKind', 88),
                        ('name', 'ชื่อที่จะกรอกลงเว็บ', 310), ('type', 'ประเภท', 85),
                        ('price', 'Price', 60), ('dur', 'ระยะเวลา', 78),
                        ('qty', 'จำนวน', 58), ('trade', 'แลกเปลี่ยน', 82)):
            self.c_tree.heading(c, text=t)
            self.c_tree.column(c, width=w, anchor='w')
        self.c_tree.tag_configure('warn', background='#3a3018', foreground='#e3b341')
        csb = ttk.Scrollbar(tw, orient='vertical', command=self.c_tree.yview)
        self.c_tree.configure(yscrollcommand=csb.set)
        self.c_tree.pack(side='left', fill='both', expand=True)
        csb.pack(side='right', fill='y')
        self.c_tree.bind('<Double-1>', lambda e: self.c_edit())
        self.c_tree.bind('<<TreeviewSelect>>', lambda e: self._c_why())

        self.c_warn = tk.Label(s2, text='', bg=C['bg'], fg=C['warn'], font=('Segoe UI', 9),
                               anchor='w', justify='left', wraplength=940)
        self.c_warn.pack(fill='x', pady=(8, 0))

        # ---------- ลงมือ ----------
        s3 = self._card(p, 'ลงมือสร้าง')
        self.cv_do = tk.BooleanVar(value=False)
        cb = tk.Checkbutton(s3, variable=self.cv_do, command=self._c_do_changed,
                            text='กดปุ่ม “สร้าง Item” จริง',
                            bg=C['bg'], fg=C['fg'], selectcolor=C['input'],
                            activebackground=C['bg'], activeforeground=C['fg'],
                            font=FB, bd=0, highlightthickness=0)
        cb.grid(row=0, column=0, sticky='w')
        self.c_mode = tk.Label(s3, text='', bg=C['bg'], fg=C['warn'], font=FM)
        self.c_mode.grid(row=0, column=1, sticky='w', padx=(10, 0))

        self.cv_hold = tk.StringVar(value=self.prefs.get('c_hold', '3'))
        tk.Label(s3, text='โหมดทดสอบ: กรอกเสร็จแล้วค้างหน้าไว้ให้ดู (วินาที)',
                 bg=C['bg'], fg=C['dim'], font=('Segoe UI', 9)
                 ).grid(row=1, column=0, sticky='w', pady=(8, 0))
        eh = self._entry(s3, width=6)
        eh.config(textvariable=self.cv_hold)
        eh.grid(row=1, column=1, sticky='w', padx=(10, 0), pady=(8, 0), ipady=3)

        run = tk.Frame(s3, bg=C['bg'])
        run.grid(row=2, column=0, columnspan=3, sticky='w', pady=(12, 0))
        self.c_btn_run = self._btn(run, '▶  เริ่มทำงาน', self.c_start, primary=True)
        self.c_btn_run.pack(side='left', ipadx=18, ipady=5)
        self.c_btn_stop = self._btn(run, '■  ยกเลิก', self.c_stop)
        self.c_btn_stop.config(state='disabled')
        self.c_btn_stop.pack(side='left', padx=(8, 0), ipadx=12, ipady=5)

        self._c_do_changed()
        self._c_refresh()

    def _c_do_changed(self):
        if self.cv_do.get():
            self.c_mode.config(text='⚠  จะสร้างไอเทมจริงบนเว็บ', fg=C['err'])
        else:
            self.c_mode.config(text='โหมดทดสอบ — กรอกฟอร์มให้ดูเฉยๆ ไม่กดสร้าง', fg=C['ok'])

    def _c_why(self):
        """คลิกแถวเหลืองในคิว แล้วบอกว่าเตือนเพราะอะไร"""
        i = self._c_sel()
        if i is None or not (0 <= i < len(self.cq)):
            return
        w = self.cq[i].get('warn') or []
        self.c_warn.config(text=('⚠  แถว #%d : %s' % (i + 1, ' · '.join(w))) if w else '',
                           fg=C['warn'])

    def _c_refresh(self):
        # คิดใหม่ทุกครั้ง — แก้ ItemKind หรือชื่อแล้วคำเตือนจะหายไปเอง
        nwarn = mark_suspects(self.cq)
        self.c_tree.delete(*self.c_tree.get_children())
        for i, d in enumerate(self.cq, 1):
            self.c_tree.insert('', 'end', tags=('warn',) if d.get('warn') else (), values=(
                i, '⚠' if d.get('warn') else '', d['kind'], d['name'], d['type'],
                d['price'] or '—', (d['dur'] + ' วัน') if d['dur'] else 'ถาวร',
                d['qty'] or '—', 'ได้' if d['trade'] else 'ไม่ได้'))
        self.c_count.config(text=('คิวว่าง' if not self.cq else f'ในคิว {len(self.cq)} ไอเทม'))
        self.c_warn.config(
            text=(f'⚠  มี {nwarn} แถวหน้าตาไม่เหมือนไอเทมทั่วไป (สีเหลือง) — '
                  'คลิกที่แถวเพื่อดูเหตุผล · ดับเบิลคลิกเพื่อแก้ · ไม่เอาก็ลบออกจากคิวได้'
                  if nwarn else ''), fg=C['warn'])

    def _c_sel(self):
        s = self.c_tree.selection()
        if not s:
            return None
        return self.c_tree.index(s[0])

    # ---------- นำเข้า / แก้ไขคิว ----------
    def c_import(self):
        if self.c_running:
            return
        path = filedialog.askopenfilename(title='เลือกไฟล์ต้นฉบับ',
                                          filetypes=[('Excel', '*.xlsx *.xlsm'), ('ทุกไฟล์', '*.*')])
        if not path:
            return
        dlg = ImportDialog(self.root, path)
        if not dlg.result:
            return
        self.save_now()
        t = self.cv_type.get().strip() or DEFAULT_TYPE
        sfx = self.cv_suffix.get().strip()
        pr = self.cv_price.get().strip()
        ml = self.cv_mail.get().strip()
        got = [row_to_item(r, t, sfx, pr, ml) for r in dlg.result]
        self.cq.extend(got)
        self._c_refresh()
        self.log(f'นำเข้าเข้าคิวสร้างไอเทม {len(got)} รายการ (รวม {len(self.cq)})', 'OK')
        log_event('create_import', count=len(got), total=len(self.cq),
                  file=os.path.basename(path))

    def c_add(self):
        if self.c_running:
            return
        d = {'kind': '', 'name': '', 'cname': '', 'desc': '',
             'type': self.cv_type.get().strip() or DEFAULT_TYPE,
             'price': self.cv_price.get().strip(), 'dur': '',
             'mail': self.cv_mail.get().strip(), 'qty': '1',
             'trade': False, 'web': True}
        if self._c_form(d, 'เพิ่มไอเทมเข้าคิว'):
            self.cq.append(d)
            self._c_refresh()

    def c_edit(self):
        if self.c_running:
            return
        i = self._c_sel()
        if i is None:
            return messagebox.showinfo('แก้ไข', 'เลือกแถวในคิวก่อนนะ')
        d = dict(self.cq[i])
        if self._c_form(d, f'แก้ไขไอเทม #{i + 1}'):
            self.cq[i] = d
            self._c_refresh()

    def c_del(self):
        i = self._c_sel()
        if i is None:
            return
        del self.cq[i]
        self._c_refresh()

    def c_clear(self):
        if self.c_running or not self.cq:
            return
        if messagebox.askyesno('ล้างคิว', f'ลบทั้ง {len(self.cq)} รายการออกจากคิว?'):
            self.cq = []
            self._c_refresh()

    def _c_form(self, d, title):
        """หน้าต่างแก้ไขไอเทม — แก้ได้ทุกช่อง คืน True ถ้ากดบันทึก"""
        top = tk.Toplevel(self.root)
        top.title(title)
        top.configure(bg=C['bg'])
        top.transient(self.root)
        top.grab_set()
        try:    # วางกลางหน้าต่างหลัก ไม่ให้ไปโผล่มุมจอ
            self.root.update_idletasks()
            x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - 640) // 2)
            y = self.root.winfo_rooty() + max(0, (self.root.winfo_height() - 520) // 2)
            top.geometry('640x520+%d+%d' % (x, y))
        except Exception:
            top.geometry('640x520')
        ok = {'v': False}

        body = tk.Frame(top, bg=C['bg'])
        body.pack(fill='both', expand=True, padx=18, pady=14)

        fields = (('ชื่อไอเทม (ที่จะกรอกลงเว็บ)', 'name', 48),
                  ('ItemKind / game_item_id', 'kind', 20),
                  ('ประเภทไอเทม', 'type', 20),
                  ('Price', 'price', 12),
                  ('ระยะเวลา (วัน) — เว้นว่าง = ถาวร', 'dur', 12),
                  ('จำนวน (แสดงบนเว็บ)', 'qty', 12),
                  ('หัวข้อจดหมาย', 'mail', 32),
                  ('คำอธิบาย', 'desc', 48))
        vs = {}
        for r, (lbl, key, w) in enumerate(fields):
            tk.Label(body, text=lbl, bg=C['bg'], fg=C['dim'],
                     font=('Segoe UI', 9)).grid(row=r, column=0, sticky='w', pady=4)
            v = tk.StringVar(value=str(d.get(key, '') or ''))
            vs[key] = v
            e = self._entry(body, width=w)
            e.config(textvariable=v)
            e.grid(row=r, column=1, sticky='w', padx=(12, 0), ipady=3)

        vt = tk.BooleanVar(value=bool(d.get('trade')))
        vw = tk.BooleanVar(value=bool(d.get('web', True)))
        for r, (txt, var) in enumerate(((' แลกเปลี่ยนได้', vt),
                                        (' เปิดใช้งานการแสดงผลบนเว็บ', vw)),
                                       start=len(fields)):
            tk.Checkbutton(body, text=txt, variable=var, bg=C['bg'], fg=C['fg'],
                           selectcolor=C['input'], activebackground=C['bg'],
                           activeforeground=C['fg'], font=FM, bd=0,
                           highlightthickness=0).grid(row=r, column=1, sticky='w',
                                                      padx=(12, 0), pady=2)

        # ตัวช่วย: ประกอบชื่อใหม่จาก "ชื่อฐาน" + คำต่อท้าย + ระยะเวลา
        hint = tk.Label(body, text='ชื่อฐาน (ตัดจำนวนชิ้นออกแล้ว): ' + (d.get('cname') or '—'),
                        bg=C['bg'], fg=C['dim'], font=('Segoe UI', 8),
                        wraplength=560, justify='left')
        hint.grid(row=len(fields) + 2, column=0, columnspan=2, sticky='w', pady=(10, 0))

        def _rebuild():
            vs['name'].set(make_item_name(d.get('cname') or vs['name'].get(),
                                          self.cv_suffix.get().strip(),
                                          vs['dur'].get().strip()))
        self._btn(body, '↻ ประกอบชื่อใหม่จากชื่อฐาน + คำต่อท้าย + ระยะเวลา', _rebuild
                  ).grid(row=len(fields) + 3, column=0, columnspan=2,
                         sticky='w', pady=(6, 0), ipadx=6, ipady=3)

        foot = tk.Frame(top, bg=C['card'], height=58)
        foot.pack(fill='x', side='bottom')
        foot.pack_propagate(False)

        def _save():
            for key, v in vs.items():
                d[key] = v.get().strip()
            d['trade'] = vt.get()
            d['web'] = vw.get()
            if not d['kind'] or not d['name']:
                return messagebox.showwarning('กรอกไม่ครบ', 'ต้องมี ItemKind และชื่อไอเทม')
            ok['v'] = True
            top.destroy()

        self._btn(foot, 'บันทึก', _save, primary=True).pack(side='right', padx=18, pady=12,
                                                            ipadx=20, ipady=4)
        self._btn(foot, 'ยกเลิก', top.destroy).pack(side='right', pady=12, ipadx=14, ipady=4)
        self.root.wait_window(top)
        return ok['v']

    # ---------- รัน ----------
    def c_stop(self):
        self.c_cancel = True
        self.log('กำลังยกเลิกการสร้างไอเทม...', 'WARN')

    def c_start(self):
        if self.c_running or self.running:
            return messagebox.showinfo('กำลังทำงาน', 'รอให้งานปัจจุบันเสร็จก่อนนะ')
        if not self.cq:
            return messagebox.showwarning('คิวว่าง', 'ยังไม่มีไอเทมในคิว — นำเข้าไฟล์ต้นฉบับก่อน')
        do = self.cv_do.get()
        nwarn = sum(1 for d in self.cq if d.get('warn'))
        if do:
            msg = f'จะสร้างไอเทมจริงบนเว็บ {len(self.cq)} รายการ\n'
            if nwarn:
                msg += (f'\n⚠  ในคิวมี {nwarn} แถวที่หน้าตาไม่เหมือนไอเทมทั่วไป (แถวสีเหลือง)\n'
                        '    ถ้ายังไม่ได้ดู แนะนำให้กดยกเลิกแล้วไล่ดูก่อน\n')
            msg += '\nตรวจชื่อ/ItemKind ในคิวเรียบร้อยแล้วใช่ไหม?'
            if not messagebox.askyesno('ยืนยัน', msg):
                return
        self.save_now()
        self.c_running = True
        self.c_cancel = False
        self.c_btn_run.config(state='disabled')
        self.c_btn_stop.config(state='normal')
        self.nb.select(self.tab_log)
        self.log('=' * 46, 'STEP')
        self.log(('เริ่มสร้างไอเทมจริง ' if do else 'เริ่มทดสอบกรอกฟอร์ม ')
                 + f'{len(self.cq)} รายการ', 'STEP')
        log_event('create_start', count=len(self.cq), commit=bool(do))
        threading.Thread(target=self._c_thread, args=(list(self.cq), do), daemon=True).start()

    def _c_thread(self, queue_rows, do):
        try:
            asyncio.run(self._c_work(queue_rows, do))
        except Exception as ex:
            log_event('error', where='create', message=str(ex)[:300])
            self.log('ผิดพลาด: ' + str(ex), 'ERR')
            self.log(traceback.format_exc(), 'ERR')
        finally:
            self.c_running = False

            def _rst():
                self.c_btn_run.config(state='normal')
                self.c_btn_stop.config(state='disabled')
                if self.results:
                    self.nb.select(self.tab_result)
            self.root.after(0, _rst)

    async def _c_work(self, rows, do):
        async with async_playwright() as pw:
            browser = await pw.chromium.launch_persistent_context(**launch_kwargs(False))
            page = browser.pages[0] if browser.pages else await browser.new_page()
            try:
                try:
                    hold = max(0, int(float(self.cv_hold.get() or 0)))
                except Exception:
                    hold = 3
                okc = errc = 0
                for i, d in enumerate(rows, 1):
                    if self.c_cancel:
                        self.log('ยกเลิกแล้ว', 'WARN')
                        break
                    self.set_progress(i - 1, len(rows), d['name'][:28])
                    self.log(f'[{i}/{len(rows)}] {d["name"]}  (ItemKind={d["kind"]})', 'STEP')
                    try:
                        if await self._c_one(page, d, do, hold):
                            okc += 1
                        else:
                            errc += 1
                    except Exception as ex:
                        errc += 1
                        self.log('   ✗ ' + str(ex)[:160], 'ERR')
                        log_event('error', where='create_one', kind_id=d.get('kind'),
                                  message=str(ex)[:200])
                self.set_progress(len(rows), len(rows), 'เสร็จ')
                self.log(f'จบ — สำเร็จ {okc} · ไม่ผ่าน {errc}', 'OK' if not errc else 'WARN')
                log_event('create_done', ok=okc, fail=errc, commit=bool(do))
            finally:
                try:
                    await browser.close()
                except Exception:
                    pass

    async def _c_fill(self, page, key, value, label=''):
        """กรอกช่องเดียว: ลอง id ก่อน ถ้าไม่เจอค่อยหาโดยอ้างข้อความข้างๆ"""
        if value is None or str(value) == '':
            return True
        sel = SEL_CREATE.get(key, '')
        r = await page.evaluate(JS_SET_VALUE, [sel, str(value)])
        if r != 'ok' and label:
            r = await page.evaluate(JS_SET_BY_LABEL, [label, str(value)])
        if r == 'ok':
            self.log(f'   · {label or key} = {str(value)[:40]}', 'INFO')
            return True
        self.log(f'   ! ไม่เจอช่อง {label or key} ({sel})', 'WARN')
        return False

    async def _c_one(self, page, d, do, hold):
        await page.goto(ITEM_CREATE_URL, wait_until='domcontentloaded', timeout=45000)
        await page.wait_for_timeout(2000)
        if any(k in page.url.lower() for k in ('login', 'signin', 'auth')):
            self.log('   ✗ ยังไม่ได้ล็อกอิน — กด “เปิดหน้า Login” ด้านบนก่อน', 'ERR')
            return False
        try:
            await page.wait_for_selector(SEL_CREATE['name'], timeout=15000)
        except Exception:
            self.log('   ! ไม่เจอฟอร์มตาม id เดิม — จะลองหาโดยอ้างชื่อช่องแทน', 'WARN')

        # [1] รายละเอียดไอเทม
        await self._c_fill(page, 'name', d['name'], 'ชื่อไอเทม')
        await self._c_fill(page, 'desc', d.get('desc', ''), 'คำอธิบาย')
        if d.get('type'):
            r = await page.evaluate(JS_PICK_TYPE, [d['type'], SEL_CREATE['type_ph']])
            if r == 'ok':
                self.log(f'   · ประเภทไอเทม = {d["type"]}', 'INFO')
            else:
                await self._c_pick_combo(page, d['type'])

        # [2] พารามิเตอร์ส่งเข้าเกม
        await self._c_fill(page, 'kind', d['kind'], 'game_item_id (ItemKind)')
        await self._c_fill(page, 'price', d.get('price', ''), 'Price')
        await self._c_fill(page, 'duration', d.get('dur', ''), 'ระยะเวลาไอเทม (วัน)')
        await self._c_fill(page, 'mail', d.get('mail', ''), 'หัวข้อจดหมาย')
        await self._c_switch(page, SEL_CREATE['trade_label'], bool(d.get('trade')))

        # [3] พารามิเตอร์แสดงบนเว็บ
        await self._c_switch(page, SEL_CREATE['web_label'], bool(d.get('web', True)))
        await page.wait_for_timeout(300)
        await self._c_fill(page, 'qty', d.get('qty', ''), 'จำนวน')

        # ตรวจว่าค่าเข้าไปจริง
        got = await page.evaluate(JS_READ_FORM, SEL_CREATE)
        bad = []
        for k, want in (('name', d['name']), ('kind', d['kind'])):
            if (got.get(k) or '').strip() != str(want).strip():
                bad.append(f'{k}: อยากได้ “{want}” แต่ในฟอร์มเป็น “{got.get(k)}”')
        if bad:
            for b in bad:
                self.log('   ! ' + b, 'WARN')

        if not do:
            self.log('   ✓ กรอกฟอร์มครบแล้ว (โหมดทดสอบ — ไม่กดสร้าง)', 'OK')
            if hold:
                await page.wait_for_timeout(hold * 1000)
            self.add_result({'kind': d['kind'], 'name': d['name'], 'id': '',
                             'type': d['type'], 'dur': d.get('dur', ''),
                             'qty': d.get('qty', '')}, 'กรอกฟอร์มแล้ว (ไม่สร้าง)')
            return not bad

        # [4] กดสร้างจริง
        btn = page.locator(f'button:has-text("{SEL_CREATE["submit"]}")').last
        if await btn.count() == 0:
            self.log(f'   ✗ ไม่เจอปุ่ม “{SEL_CREATE["submit"]}”', 'ERR')
            return False
        try:
            async with page.expect_response(
                    lambda r: r.request.method in ('POST', 'PUT', 'PATCH')
                    and 'item' in r.url.lower(), timeout=20000) as ri:
                await btn.click(timeout=10000)
            resp = await ri.value
            code = resp.status
        except Exception:
            code = 0
        await page.wait_for_timeout(1500)
        good = (code == 0) or (200 <= code < 300)
        self.log('   ' + ('✓ สร้างแล้ว' if good else f'✗ เว็บตอบ HTTP {code}')
                 + (f' (HTTP {code})' if good and code else ''),
                 'OK' if good else 'ERR')
        self.add_result({'kind': d['kind'], 'name': d['name'], 'id': '',
                         'type': d['type'], 'dur': d.get('dur', ''),
                         'qty': d.get('qty', '')},
                        'สร้างแล้ว' if good else f'ไม่สำเร็จ (HTTP {code})')
        log_event('create_item', kind_id=d['kind'], name=d['name'],
                  ok=bool(good), http=code)
        return good

    async def _c_switch(self, page, label, want):
        r = await page.evaluate(JS_SET_SWITCH, [label, bool(want)])
        if r == 'ok':
            self.log(f'   · {label} = {"เปิด" if want else "ปิด"}', 'INFO')
        else:
            self.log(f'   ! ไม่เจอสวิตช์ “{label}”', 'WARN')

    async def _c_pick_combo(self, page, value):
        """combobox แบบ React — คลิกเปิดแล้วคลิกตัวเลือก"""
        for trig in ('[role="combobox"]', 'text="%s"' % SEL_CREATE['type_ph']):
            try:
                t = page.locator(trig).first
                if await t.count() == 0:
                    continue
                await t.click(timeout=4000)
                await page.wait_for_timeout(400)
                opt = page.locator(f'[role="option"]:has-text("{value}")').first
                if await opt.count() == 0:
                    opt = page.locator(f'text="{value}"').last
                if await opt.count() > 0:
                    await opt.click(timeout=4000)
                    self.log(f'   · ประเภทไอเทม = {value}', 'INFO')
                    return True
            except Exception:
                continue
        self.log(f'   ! เลือกประเภทไอเทม “{value}” ไม่ได้', 'WARN')
        return False

    def _build_result(self):
        bar = tk.Frame(self.tab_result, bg=C['bg'])
        bar.pack(fill='x', padx=14, pady=12)
        self._btn(bar, '📋  คัดลอก ID ทั้งหมด', self.copy_ids).pack(side='left', ipadx=10, ipady=4)
        self._btn(bar, '⬇  Export Excel', lambda: self.export('xlsx')).pack(side='left', padx=8, ipadx=10, ipady=4)
        self._btn(bar, '⬇  Export CSV', lambda: self.export('csv')).pack(side='left', ipadx=10, ipady=4)
        self._btn(bar, '🗑  ล้าง', self.clear_results).pack(side='right', ipadx=10, ipady=4)
        self.lbl_count = tk.Label(bar, text='พบ 0 รายการ', bg=C['bg'], fg=C['dim'], font=FB)
        self.lbl_count.pack(side='right', padx=14)

        wrap = tk.Frame(self.tab_result, bg=C['bg'])
        wrap.pack(fill='both', expand=True, padx=14, pady=(0, 14))
        style = ttk.Style()
        style.configure('TR.Treeview', background=C['card'], fieldbackground=C['card'],
                        foreground=C['fg'], rowheight=26, borderwidth=0, font=('Segoe UI', 9))
        style.configure('TR.Treeview.Heading', background=C['input'], foreground=C['dim'],
                        font=('Segoe UI', 9, 'bold'), borderwidth=0)
        cols = ('id', 'name', 'type', 'kind', 'notes')
        self.tree = ttk.Treeview(wrap, columns=cols, show='headings', style='TR.Treeview')
        for c, t, w in (('id', 'Aztek Item Id', 100), ('name', 'ชื่อ', 380),
                        ('type', 'ประเภท', 90), ('kind', 'ItemKind', 100), ('notes', 'หมายเหตุ', 220)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor='w')
        sb = ttk.Scrollbar(wrap, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

    # ------------------------------------------------------------------
    #  แท็บตรวจระบบ
    # ------------------------------------------------------------------
    def _build_check(self):
        p = self.tab_check

        # ---- แถบตั้งค่าที่เก็บ log ----
        box = tk.LabelFrame(p, text='  ที่เก็บบันทึกการใช้งาน  ', bg=C['bg'], fg=C['dim'],
                            font=('Segoe UI', 9, 'bold'), bd=1, relief='solid')
        box.pack(fill='x', padx=14, pady=(12, 0))
        inner = tk.Frame(box, bg=C['bg'])
        inner.pack(fill='x', padx=12, pady=10)

        r1 = tk.Frame(inner, bg=C['bg'])
        r1.pack(fill='x')
        tk.Label(r1, text='บันทึกในเครื่อง', bg=C['bg'], fg=C['dim'], font=('Segoe UI', 9),
                 width=16, anchor='w').pack(side='left')
        self.lbl_local = tk.Label(r1, text='', bg=C['bg'], fg=C['fg'],
                                  font=('Consolas', 9), anchor='w')
        self.lbl_local.pack(side='left')
        self._btn(r1, '📂  เปิด', self.open_log_dir).pack(side='right', padx=(8, 0), ipadx=8, ipady=2)
        self._btn(r1, '↺  คืนค่าเดิม', self.reset_local).pack(side='right', padx=(8, 0), ipadx=8, ipady=2)
        self._btn(r1, '📁  เลือกโฟลเดอร์', self.pick_local).pack(side='right', ipadx=8, ipady=2)

        r2 = tk.Frame(inner, bg=C['bg'])
        r2.pack(fill='x', pady=(8, 0))
        tk.Label(r2, text='โฟลเดอร์กลางของทีม', bg=C['bg'], fg=C['dim'], font=('Segoe UI', 9),
                 width=16, anchor='w').pack(side='left')
        self.lbl_central = tk.Label(r2, text='', bg=C['bg'], fg=C['fg'],
                                    font=('Consolas', 9), anchor='w')
        self.lbl_central.pack(side='left')
        self._btn(r2, '✕  ล้าง', self.clear_central).pack(side='right', padx=(8, 0), ipadx=8, ipady=2)
        self._btn(r2, '📁  เลือกโฟลเดอร์', self.pick_central).pack(side='right', ipadx=8, ipady=2)

        r3 = tk.Frame(inner, bg=C['bg'])
        r3.pack(fill='x', pady=(8, 0))
        tk.Label(r3, text='โฟลเดอร์โปรแกรม', bg=C['bg'], fg=C['dim'], font=('Segoe UI', 9),
                 width=16, anchor='w').pack(side='left')
        tk.Label(r3, text=DATA_DIR + '   (เก็บ login ของ Chrome ไว้ที่นี่ ย้ายไม่ได้)',
                 bg=C['bg'], fg=C['dim'], font=('Consolas', 9), anchor='w').pack(side='left')
        self._btn(r3, '📂  เปิด', self.open_data_dir).pack(side='right', ipadx=8, ipady=2)

        self.lbl_todo = tk.Label(inner, text='', bg=C['bg'], fg=C['warn'],
                                 font=('Segoe UI', 9), wraplength=980, justify='left', anchor='w')
        self.lbl_todo.pack(fill='x', pady=(9, 0))
        self._refresh_central()

        bar = tk.Frame(p, bg=C['bg'])
        bar.pack(fill='x', padx=14, pady=12)
        self.btn_chk_logic = self._btn(bar, '⚡  ตรวจตรรกะ (เร็ว ไม่ต้องต่อเน็ต)', self.run_check_logic)
        self.btn_chk_logic.pack(side='left', ipadx=10, ipady=5)
        self.btn_chk_full = self._btn(bar, '🌐  ตรวจเต็ม (เปิดเว็บจริง)', self.run_check_full, primary=True)
        self.btn_chk_full.pack(side='left', padx=8, ipadx=10, ipady=4)
        self._btn(bar, '📋  คัดลอกผล', self.copy_check).pack(side='right', ipadx=10, ipady=5)
        self.lbl_chk = tk.Label(bar, text='', bg=C['bg'], fg=C['dim'], font=FB)
        self.lbl_chk.pack(side='right', padx=14)

        tk.Label(p, text='ใช้ตรวจว่าเครื่องมือยังทำงานถูกต้องไหม — "ตรวจเต็ม" จะเปิดเว็บจริงแล้วไล่เช็กทีละจุด '
                         'ถ้าเว็บเปลี่ยนหน้าตา จะบอกได้เลยว่าพังตรงไหนและต้องแก้ค่าตัวไหน',
                 bg=C['bg'], fg=C['dim'], font=('Segoe UI', 9), wraplength=1000,
                 justify='left').pack(anchor='w', padx=14)

        wrap = tk.Frame(p, bg=C['bg'])
        wrap.pack(fill='both', expand=True, padx=14, pady=(10, 14))
        cols = ('st', 'name', 'detail')
        self.chk_tree = ttk.Treeview(wrap, columns=cols, show='headings', style='TR.Treeview')
        for c, t, w in (('st', 'ผล', 60), ('name', 'รายการตรวจ', 420), ('detail', 'รายละเอียด', 480)):
            self.chk_tree.heading(c, text=t)
            self.chk_tree.column(c, width=w, anchor='w')
        self.chk_tree.tag_configure('ok', foreground=C['ok'])
        self.chk_tree.tag_configure('bad', foreground=C['err'])
        sb = ttk.Scrollbar(wrap, orient='vertical', command=self.chk_tree.yview)
        self.chk_tree.configure(yscrollcommand=sb.set)
        self.chk_tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')
        self.check_rows = []

    def open_log_dir(self):
        try:
            os.startfile(LOCAL_LOG_DIR)
        except Exception:
            messagebox.showinfo('ที่เก็บ log', LOCAL_LOG_DIR)

    def pick_local(self):
        path = filedialog.askdirectory(title='เลือกที่เก็บบันทึกการใช้งานในเครื่องนี้')
        if not path:
            return
        moved, err = set_local_log_dir(path)
        self._refresh_central()
        if err:
            messagebox.showerror('ย้ายที่เก็บ log', 'ย้ายไม่สำเร็จ: ' + err)
            return
        log_event('set_local_log_dir', path=path, moved=moved)
        self.log(f'ย้ายที่เก็บ log ไปที่ {path} (ย้ายไฟล์เดิม {moved} ไฟล์)', 'OK')
        messagebox.showinfo('ย้ายที่เก็บ log',
                            f'เรียบร้อย\n\nที่ใหม่: {path}\nย้ายประวัติเดิมตามไปให้ {moved} ไฟล์')

    def reset_local(self):
        if os.path.abspath(LOCAL_LOG_DIR) == os.path.abspath(DATA_DIR):
            return
        moved, err = set_local_log_dir(DATA_DIR)
        self._refresh_central()
        self.log(f'คืนค่าที่เก็บ log กลับเป็นค่าเริ่มต้น (ย้ายไฟล์ {moved})', 'OK')

    def _refresh_central(self):
        default = os.path.abspath(LOCAL_LOG_DIR) == os.path.abspath(DATA_DIR)
        self.lbl_local.config(text=LOCAL_LOG_DIR + ('   (ค่าเริ่มต้น)' if default else ''),
                              fg=C['dim'] if default else C['fg'])
        self.lbl_central.config(
            text=CENTRAL_LOG_DIR or '— ยังไม่ได้ตั้ง (เก็บเฉพาะในเครื่องนี้) —',
            fg=C['fg'] if CENTRAL_LOG_DIR else C['dim'])
        todo = pending_items()
        self.lbl_todo.config(text=('⚠  ' + todo[0]) if todo else
                             '✓  ตั้งค่าครบแล้ว — log จะถูกเขียนทั้งในเครื่องและโฟลเดอร์กลาง',
                             fg=C['warn'] if todo else C['ok'])

    def open_data_dir(self):
        try:
            os.startfile(DATA_DIR)
        except Exception:
            messagebox.showinfo('ที่เก็บข้อมูล', DATA_DIR)

    def pick_central(self):
        path = filedialog.askdirectory(title='เลือกโฟลเดอร์กลางของทีม (ที่ทุกคนเข้าถึงได้)')
        if not path:
            return
        set_central_dir(path)
        self._refresh_central()
        log_event('set_central_dir', path=path)
        self.log(f'ตั้งโฟลเดอร์กลางเป็น {path}', 'OK')
        messagebox.showinfo('โฟลเดอร์กลาง',
                            'ตั้งค่าเรียบร้อย\n\nจากนี้ log จะถูกเขียนทั้งในเครื่องและในโฟลเดอร์นี้\n'
                            'ให้คนอื่นในทีมตั้งโฟลเดอร์เดียวกัน แล้วจะรวมกันเอง')

    def clear_central(self):
        if not CENTRAL_LOG_DIR:
            return
        set_central_dir('')
        self._refresh_central()
        self.log('ล้างโฟลเดอร์กลางแล้ว — กลับไปเก็บเฉพาะในเครื่อง', 'WARN')

    def _show_check(self, results, append=False):
        if not append:
            self.chk_tree.delete(*self.chk_tree.get_children())
            self.check_rows = []
        self.check_rows.extend(results)
        for x in results:
            detail = x['detail']
            if not x['ok'] and x['fix']:
                detail = (detail + '   →  ' + x['fix']).strip(' →')
            self.chk_tree.insert('', 'end', values=('ผ่าน' if x['ok'] else 'พัง', x['name'], detail),
                                 tags=('ok' if x['ok'] else 'bad',))
        bad = sum(1 for x in self.check_rows if not x['ok'])
        total = len(self.check_rows)
        self.lbl_chk.config(text=f'ผ่าน {total - bad}/{total}' + ('  ✓ ปกติดี' if not bad else f'  ✗ พัง {bad} จุด'),
                            fg=C['ok'] if not bad else C['err'])

    def copy_check(self):
        if not self.check_rows:
            return messagebox.showinfo('ตรวจระบบ', 'ยังไม่ได้ตรวจ')
        lines = [f'TR Plus Ultra v{APP_VERSION} — ผลตรวจระบบ {datetime.now():%Y-%m-%d %H:%M}', '']
        for x in self.check_rows:
            mark = 'OK  ' if x['ok'] else 'FAIL'
            lines.append(f'[{mark}] {x["name"]}  |  {x["detail"]}'
                         + (f'  ->  {x["fix"]}' if (not x['ok'] and x['fix']) else ''))
        bad = sum(1 for x in self.check_rows if not x['ok'])
        lines += ['', f'สรุป: ผ่าน {len(self.check_rows) - bad}/{len(self.check_rows)}']
        self.root.clipboard_clear()
        self.root.clipboard_append('\n'.join(lines))
        self.log('คัดลอกผลตรวจแล้ว ส่งให้คนดูแลเครื่องมือได้เลย', 'OK')

    def run_check_logic(self):
        self.nb.select(self.tab_check)
        try:
            res = run_logic_tests()
            self._show_check(res)
            bad = sum(1 for x in res if not x['ok'])
            log_event('selfcheck', mode='logic', total=len(res), failed=bad,
                      failed_items=[x['name'] for x in res if not x['ok']][:20])
            self.log('ตรวจตรรกะเสร็จ', 'OK')
        except Exception as ex:
            messagebox.showerror('ตรวจระบบ', str(ex))

    def run_check_full(self):
        if self.running:
            return messagebox.showinfo('กำลังทำงาน', 'รอให้งานปัจจุบันเสร็จก่อนนะ')
        self.nb.select(self.tab_check)
        self._show_check(run_logic_tests())
        self.running = True
        self.btn_chk_full.config(state='disabled', text='กำลังตรวจ...')
        threading.Thread(target=self._check_thread, daemon=True).start()

    def _check_thread(self):
        try:
            asyncio.run(self._check_web())
        except Exception as ex:
            msg = str(ex)[:120]
            self.root.after(0, lambda m=msg: self._show_check(
                [_t(False, 'เปิดเบราว์เซอร์ได้', m)], append=True))
        finally:
            self.running = False
            self.root.after(0, lambda: self.btn_chk_full.config(
                state='normal', text='🌐  ตรวจเต็ม (เปิดเว็บจริง)'))

    async def _check_web(self):
        self.log('เริ่มตรวจเว็บจริง...', 'STEP')
        async with async_playwright() as pw:
            browser = await pw.chromium.launch_persistent_context(**launch_kwargs(True))
            page = browser.pages[0] if browser.pages else await browser.new_page()
            try:
                res = await run_web_tests(page, log=lambda m: self.log(m))
            finally:
                try:
                    await browser.close()
                except Exception:
                    pass
        self.root.after(0, lambda: self._show_check(res, append=True))
        bad = sum(1 for x in res if not x['ok'])
        log_event('selfcheck', mode='web', total=len(res), failed=bad,
                  failed_items=[x['name'] for x in res if not x['ok']][:20])
        self.log(f'ตรวจเว็บเสร็จ — พัง {bad} จุด' if bad else 'ตรวจเว็บเสร็จ — ปกติดีทุกจุด',
                 'WARN' if bad else 'OK')

    def _build_log(self):
        self.log_box = scrolledtext.ScrolledText(
            self.tab_log, bg=C['input'], fg=C['fg'], bd=0, font=('Consolas', 9),
            insertbackground=C['fg'], wrap='word')
        self.log_box.pack(fill='both', expand=True, padx=14, pady=12)
        for tag, col in (('INFO', C['dim']), ('STEP', C['accent']), ('OK', C['ok']),
                         ('WARN', C['warn']), ('ERR', C['err'])):
            self.log_box.tag_config(tag, foreground=col)

    # ---------- utility ----------
    def log(self, msg, kind='INFO'):
        def _do():
            self.log_box.insert('end', f'[{datetime.now():%H:%M:%S}] {msg}\n', kind)
            self.log_box.see('end')
        self.root.after(0, _do)

    def set_progress(self, done, total, note=''):
        def _do():
            self.progress['value'] = (done / total * 100) if total else 0
            self.lbl_stat.config(text=f'{done}/{total}  {note}' if total else note)
        self.root.after(0, _do)

    def add_result(self, row, notes=''):
        if row.get('id') and any(r['id'] == row['id'] for r in self.results):
            return
        row = dict(row)
        row['notes'] = notes
        self.results.append(row)

        def _do():
            self.tree.insert('', 'end', values=(row['id'], row['name'], row['type'],
                                                row['kind'], row['notes']))
            self.lbl_count.config(text=f'พบ {len(self.results)} รายการ')
        self.root.after(0, _do)

    def clear_results(self):
        self.results.clear()
        self.not_found.clear()
        self.tree.delete(*self.tree.get_children())
        self.lbl_count.config(text='พบ 0 รายการ')

    def copy_ids(self):
        if not self.results:
            return messagebox.showinfo('ผลลัพธ์', 'ยังไม่มีผลลัพธ์')
        ids = ', '.join(r['id'] for r in self.results)
        self.root.clipboard_clear()
        self.root.clipboard_append(ids)
        self.log(f'คัดลอก {len(self.results)} ID แล้ว', 'OK')

    def export(self, kind):
        if not self.results:
            return messagebox.showinfo('ผลลัพธ์', 'ยังไม่มีผลลัพธ์')
        ext = '.xlsx' if kind == 'xlsx' else '.csv'
        path = filedialog.asksaveasfilename(
            defaultextension=ext, filetypes=[(kind.upper(), '*' + ext)],
            initialfile=f'TR_Items_{datetime.now():%Y%m%d_%H%M}{ext}')
        if not path:
            return
        header = ['Aztek Item Id', 'ชื่อ', 'ประเภท', 'ItemKind', 'หมายเหตุ']
        rows = [[r['id'], r['name'], r['type'], r['kind'], r['notes']] for r in self.results]
        if kind == 'xlsx' and XLSX_OK:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = 'Results'
            ws.append(header)
            for c in ws[1]:
                c.font = Font(bold=True, color='FFFFFF')
                c.fill = PatternFill('solid', fgColor='1F6FEB')
            for r in rows:
                ws.append(r)
            for i, w in enumerate([14, 46, 12, 12, 26], 1):
                ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
            wb.save(path)
        else:
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                w = csv.writer(f)
                w.writerow(header)
                w.writerows(rows)
        self.log(f'บันทึกไฟล์แล้ว: {path}', 'OK')
        log_event('export', format=kind, rows=len(rows), file=os.path.basename(path))

    def download_template(self):
        path = filedialog.asksaveasfilename(
            defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx'), ('CSV', '*.csv')],
            initialfile='TR_Item_Template.xlsx')
        if not path:
            return
        try:
            build_template(path)
            self.log(f'สร้าง template แล้ว: {path}', 'OK')
        except Exception as ex:
            messagebox.showerror('Template', str(ex))

    # ---------- รายการที่จะค้นหา ----------
    def toggle_imp(self):
        self.imp_open = not self.imp_open
        if self.imp_open:
            # กางแล้วต้องเห็นจริง — ถ้าหน้าต่างเตี้ยไป ขยายให้เอง
            try:
                self.root.update_idletasks()
                want = min(950, self.root.winfo_screenheight() - 80)
                if self.root.winfo_height() < want:
                    self.root.geometry('%dx%d' % (max(1080, self.root.winfo_width()), want))
            except Exception:
                pass
            self.imp_box.pack(fill='both', expand=True, padx=14, pady=(12, 0),
                              before=self.s_tail)
            self.btn_imp.config(text='▾  ซ่อนรายการที่จะค้นหา')
        else:
            self.imp_box.pack_forget()
            self.btn_imp.config(text='▸  ดูรายการที่จะค้นหา')

    def _imp_why(self):
        s = self.imp_tree.selection()
        if not s:
            return
        i = self.imp_tree.index(s[0])
        w = self.imported[i].get('warn') if 0 <= i < len(self.imported) else None
        self.imp_why.config(text=('⚠  แถว #%d : %s' % (i + 1, ' · '.join(w))) if w else '')

    def refresh_imp(self):
        """วาดรายการที่จะค้นหาใหม่ — เรียกทุกครั้งที่ self.imported เปลี่ยน"""
        nwarn = mark_suspects(self.imported)
        self.imp_tree.delete(*self.imp_tree.get_children())
        for i, r in enumerate(self.imported, 1):
            dur = r.get('dur', '')
            dur = 'ถาวร' if dur == '' else ('— ไม่กรอง —' if dur == 'any' else str(dur) + ' วัน')
            trade = {'yes': 'ได้', 'no': 'ไม่ได้', 'any': '— ไม่กรอง —'}.get(
                r.get('trade', 'any'), r.get('trade', ''))
            name = r.get('disp') or r.get('name') or ''
            self.imp_tree.insert('', 'end', tags=('warn',) if r.get('warn') else (),
                                 values=(i, '⚠' if r.get('warn') else '',
                                         r.get('kind', ''), name, dur, trade,
                                         r.get('qty') or '—'))
        n = len(self.imported)
        self.imp_cnt.config(text=f'{n} รายการ' + (f'  ·  น่าสงสัย {nwarn}' if nwarn else ''))
        self.btn_imp.config(state='normal' if n else 'disabled')
        self.btn_multi.config(state='normal' if n else 'disabled')
        self.lbl_imp.config(text=(f'⚠  มี {nwarn} รายการหน้าตาไม่เหมือนไอเทมทั่วไป — กางดูได้'
                                  if nwarn else ''))
        self.imp_why.config(text='')

    def imp_del(self):
        sel = sorted((self.imp_tree.index(s) for s in self.imp_tree.selection()), reverse=True)
        if not sel:
            return messagebox.showinfo('รายการ', 'เลือกแถวที่จะเอาออกก่อนนะ')
        for i in sel:
            if 0 <= i < len(self.imported):
                del self.imported[i]
        self.refresh_imp()
        self.lbl_file.config(text=f'{self.imp_src} — {len(self.imported)} รายการ')
        self.log(f'เอาออกจากรายการ {len(sel)} ตัว (เหลือ {len(self.imported)})', 'INFO')

    def imp_copy(self):
        if not self.imported:
            return
        ids = ', '.join(str(r.get('kind', '')) for r in self.imported if r.get('kind'))
        self.root.clipboard_clear()
        self.root.clipboard_append(ids)
        self.log(f'คัดลอก {len(self.imported)} เลขไอเทมแล้ว', 'OK')

    def pick_file(self):
        path = filedialog.askopenfilename(filetypes=[('Excel / CSV', '*.xlsx *.xlsm *.csv')])
        if not path:
            return
        try:
            if path.lower().endswith('.csv'):
                rows = read_template(path)
                src = os.path.basename(path)
            else:
                if not XLSX_OK:
                    raise RuntimeError('เครื่องนี้อ่านไฟล์ Excel ไม่ได้ ลองใช้ .csv แทน')
                dlg = ImportDialog(self.root, path)
                if dlg.result is None:
                    return
                rows = dlg.result
                names = getattr(dlg, 'sheet_names', None) or (
                    [dlg.sheet_name] if dlg.sheet_name else [])
                if len(names) > 3:
                    where = '%d ชีท' % len(names)
                else:
                    where = ', '.join(names) or '-'
                src = f'{os.path.basename(path)}  ›  {where}'
            if not rows:
                messagebox.showwarning('นำเข้า', 'ไม่พบรายการที่ใช้ได้ในไฟล์นี้')
                return
            self.imported = rows
            self.imp_src = src
            self.lbl_file.config(text=f'{src} — {len(rows)} รายการ')
            self.refresh_imp()
            self.log(f'นำเข้าจาก {src}: {len(rows)} รายการ', 'OK')
            log_event('import', file=os.path.basename(path), source=src, rows=len(rows))
        except Exception as ex:
            log_event('error', where='import', message=str(ex)[:300],
                      file=os.path.basename(path))
            messagebox.showerror('อ่านไฟล์ไม่ได้', str(ex))

    def save_now(self):
        self.prefs.update({
            'deep': self.v_deep.get(), 'dur': self.v_dur.get().strip(),
            'trade': self.v_trade.get(), 'qty': self.v_qty.get().strip(),
            'headless': self.v_headless.get(),
            'c_type': self.cv_type.get().strip(),
            'c_suffix': self.cv_suffix.get().strip(),
            'c_price': self.cv_price.get().strip(),
            'c_mail': self.cv_mail.get().strip(),
            'c_hold': self.cv_hold.get().strip(),
        })
        save_prefs(self.prefs)

    def deep_criteria(self):
        if not self.v_deep.get():
            return {'dur': 'any', 'trade': 'any', 'qty': ''}
        return {'dur': dur_from_text(self.v_dur.get()), 'trade': self.v_trade.get(),
                'qty': self.v_qty.get().strip()}

    def do_cancel(self):
        self.cancel = True
        self.log('กำลังยกเลิก...', 'WARN')

    # ---------- login ----------
    def open_login(self):
        if self.running:
            return messagebox.showinfo('กำลังทำงาน', 'รอให้งานปัจจุบันเสร็จก่อนนะ')
        log_event('login_open')
        threading.Thread(target=lambda: asyncio.run(self._login()), daemon=True).start()

    async def _login(self):
        self.log('เปิด Chrome ให้ล็อกอิน — ล็อกอินเสร็จแล้วปิดหน้าต่างได้เลย', 'STEP')
        try:
            async with async_playwright() as pw:
                b = await pw.chromium.launch_persistent_context(**launch_kwargs(False))
                page = b.pages[0] if b.pages else await b.new_page()
                await page.goto(ITEM_LIST_URL, wait_until='domcontentloaded', timeout=60000)
                while True:
                    await asyncio.sleep(1)
                    if not b.pages:
                        break
        except Exception as ex:
            self.log('ปิดหน้าต่าง login แล้ว (' + str(ex)[:80] + ')', 'INFO')
        self.log('พร้อมใช้งาน', 'OK')

    # ---------- run ----------
    def run_single(self):
        kind = self.v_kind.get().strip()
        name = self.v_name.get().strip()
        if not kind and not name:
            return messagebox.showwarning('ค้นหา', 'กรอก ItemKind หรือชื่อไอเทมอย่างน้อยหนึ่งช่อง')
        c = {'kind': kind, 'name': name}
        c.update(self.deep_criteria())
        self.start([c])

    def run_multi(self):
        if not self.imported:
            return
        if any(has_deep(c) for c in self.imported) and not self.v_deep.get():
            self.log('พบค่า deep check ในไฟล์ → เปิด Deep Check อัตโนมัติ', 'INFO')
        self.start([dict(c) for c in self.imported])

    def start(self, criteria):
        if self.running:
            return
        self.save_now()
        self.running = True
        self.cancel = False
        self.clear_results()
        self.btn_single.config(state='disabled')
        self.btn_multi.config(state='disabled')
        self.btn_cancel.config(state='normal')
        self.nb.select(self.tab_log)
        self.log('=' * 46, 'STEP')
        self.log(f'เริ่มค้นหา {len(criteria)} รายการ', 'STEP')
        self._run_started = datetime.now()
        d = self.deep_criteria()
        log_event('search_start', count=len(criteria), deep=bool(self.v_deep.get()),
                  dur=d['dur'], trade=d['trade'], qty=d['qty'],
                  kinds=[c.get('kind') for c in criteria][:50])
        threading.Thread(target=self._thread, args=(criteria,), daemon=True).start()

    def _thread(self, criteria):
        try:
            asyncio.run(self._work(criteria))
        except Exception as ex:
            log_event('error', where='search', message=str(ex)[:300])
            self.log('ผิดพลาด: ' + str(ex), 'ERR')
            self.log(traceback.format_exc(), 'ERR')
        finally:
            self.running = False

            def _rst():
                self.btn_single.config(state='normal')
                self.btn_multi.config(state='normal' if self.imported else 'disabled')
                self.btn_cancel.config(state='disabled')
                if self.results:
                    self.nb.select(self.tab_result)
            self.root.after(0, _rst)

    # ---------- automation ----------
    async def _work(self, criteria):
        async with async_playwright() as pw:
            if not find_chrome_exe():
                self.log('ไม่เจอ Chrome ในเครื่อง — ใช้ Chromium ของ Playwright แทน', 'WARN')
            browser = await pw.chromium.launch_persistent_context(
                **launch_kwargs(self.v_headless.get()))
            page = browser.pages[0] if browser.pages else await browser.new_page()
            try:
                await self._search_all(page, criteria)
            finally:
                try:
                    await browser.close()
                except Exception:
                    pass

    async def _goto_list(self, page):
        if not page.url.rstrip('/').endswith('/shop/items'):
            await page.goto(ITEM_LIST_URL, wait_until='domcontentloaded', timeout=45000)
            await page.wait_for_timeout(1500)

    async def _apply_search(self, page, keyword, dur_filter=None):
        await page.wait_for_selector(SEL['search_input'], timeout=20000)
        await page.fill(SEL['search_input'], keyword or '')
        f = page.locator(SEL['filter_duration'])
        if await f.count() > 0:
            await f.first.fill('' if dur_filter is None else str(dur_filter))
        await page.wait_for_timeout(250)
        await page.locator(SEL['search_button']).first.click()
        await page.wait_for_timeout(1400)

    async def _read_all_pages(self, page):
        all_rows, seen, pnum = [], set(), 1
        total = await page.evaluate(JS_TOTAL_PAGES)
        while True:
            if self.cancel:
                break
            rows = await page.evaluate(JS_READ_ROWS, COL)
            fresh = [r for r in rows if r['id'] not in seen]
            for r in fresh:
                seen.add(r['id'])
            all_rows.extend(fresh)
            self.log(f'  หน้า {pnum}/{total}: {len(rows)} แถว (ใหม่ {len(fresh)})')
            if pnum >= total or pnum > 200:
                break
            first = rows[0]['id'] if rows else None
            nxt = page.locator(f'button:has-text("{SEL["next_text"]}")')
            if await nxt.count() == 0 or await nxt.first.is_disabled():
                break
            await nxt.first.click()
            pnum += 1
            await page.wait_for_timeout(900)
            try:
                await page.wait_for_function(
                    """(prev) => {
                        const td = document.querySelector('table tbody tr td');
                        return td && (td.innerText || '').trim() !== prev;
                    }""", arg=first, timeout=8000)
            except Exception:
                break
        return all_rows

    async def _open_detail(self, page, item_id):
        ok = await page.evaluate(JS_CLICK_ROW, str(item_id))
        if not ok:
            await self._apply_search(page, str(item_id))
            ok = await page.evaluate(JS_CLICK_ROW, str(item_id))
            if not ok:
                raise RuntimeError(f'หาแถวของ ID {item_id} ไม่เจอ')
        await page.wait_for_selector(SEL['d_name'], timeout=20000)
        await page.wait_for_timeout(400)

    async def _read_detail(self, page):
        return await page.evaluate(JS_READ_DETAIL, {
            'name': SEL['d_name'], 'kind': SEL['d_kind'], 'price': SEL['d_price'],
            'duration': SEL['d_duration'], 'qty': SEL['d_qty'],
            'tradeLabel': SEL['d_trade_label'],
        })

    async def _search_all(self, page, criteria):
        await self._goto_list(page)
        for i, c in enumerate(criteria):
            if self.cancel:
                break
            shown = c.get('disp') or c.get('name') or ''
            label = f"#{i + 1} Kind={c.get('kind') or '-'}" + (f"  {shown}" if shown else '')
            self.log(f'── {label}', 'STEP')

            use_filter = bool(duration_only_numeric(c))
            await self._goto_list(page)
            await self._apply_search(page, c.get('kind') or c.get('name'),
                                     c['dur'] if use_filter else None)
            if use_filter:
                self.log(f"  ใช้ฟิลเตอร์ DurationIndex={c['dur']} (ข้าม deep check)")

            rows = await self._read_all_pages(page)
            if c.get('name'):
                before = len(rows)
                want = norm_name(c['name'])
                rows = [r for r in rows if want in norm_name(r['name'])]
                self.log(f"  กรองชื่อ {c['name']!r}: {before} → {len(rows)}")
            self.log(f'  พบ {len(rows)} รายการ', 'INFO' if rows else 'WARN')
            if not rows:
                self.not_found.append((label, 'ไม่พบแถวที่ตรงเงื่อนไข'))
                continue

            if not (has_deep(c) and not use_filter):
                for r in rows:
                    self.add_result(r, f"{c['dur']} วัน" if use_filter else '')
                log_event('search_item', kind=c.get('kind'), name=c.get('disp') or c.get('name'),
                          found=len(rows), passed=len(rows), deep=False,
                          ids=[r['id'] for r in rows][:20])
                self.set_progress(i + 1, len(criteria), 'เสร็จ')
                continue

            self.log(f'  Deep check {len(rows)} รายการ...', 'STEP')
            passed = 0
            for k, r in enumerate(rows):
                if self.cancel:
                    break
                self.set_progress(k + 1, len(rows), f"deep: {r['id']}")
                try:
                    await self._open_detail(page, r['id'])
                    d = await self._read_detail(page)
                    ok, why = match_deep(d, c)
                    if ok:
                        self.add_result(r, why)
                        passed += 1
                        self.log(f"  ✓ {r['id']}  {r['name']}  [{why}]", 'OK')
                    else:
                        self.log(f"  ✗ {r['id']}  {r['name']}  ({why})")
                except Exception as ex:
                    self.log(f"  ! {r['id']} ตรวจไม่ได้: {ex}", 'WARN')
                await page.go_back(wait_until='domcontentloaded')
                await page.wait_for_timeout(500)
            log_event('search_item', kind=c.get('kind'), name=c.get('disp') or c.get('name'),
                      found=len(rows), passed=passed, deep=True,
                      want_dur=c.get('dur'), want_trade=c.get('trade'), want_qty=c.get('qty'),
                      ids=[r['id'] for r in self.results][-passed:][:20] if passed else [])
            if not passed and not self.cancel:
                self.not_found.append((label, f'เจอ {len(rows)} แต่ไม่ผ่าน deep check'))
            self.set_progress(i + 1, len(criteria), 'เสร็จ')

        self.log('=' * 46, 'STEP')
        if self.cancel:
            self.log('ยกเลิกโดยผู้ใช้', 'WARN')
        self.log(f'รวมทั้งหมด {len(self.results)} รายการ', 'OK' if self.results else 'WARN')
        if self.results:
            self.log('IDs: ' + ', '.join(r['id'] for r in self.results), 'OK')
        for lb, why in self.not_found:
            self.log(f'  ✗ {lb}  ({why})', 'WARN')
        secs = 0
        try:
            secs = round((datetime.now() - self._run_started).total_seconds(), 1)
        except Exception:
            pass
        log_event('search_done', criteria=len(criteria), found=len(self.results),
                  not_found=len(self.not_found), seconds=secs, cancelled=bool(self.cancel),
                  ids=[r['id'] for r in self.results][:200])

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    App().run()

# -*- coding: utf-8 -*-
# TR_APP_MARKER — ห้ามลบบรรทัดนี้ ตัวเปิดใช้เช็กว่าโหลดไฟล์ถูกตัว
"""
TR Plus Ultra — โปรแกรมหลัก
===========================
V0.1 : ค้นหาไอเทมบนหลังบ้าน HoF เว็บใหม่ (aztek-tools-v2)

ไฟล์นี้อยู่บน GitHub ตัวเปิด (.exe) จะโหลดมารันทุกครั้ง
แก้ไฟล์นี้แล้ว push = ทุกคนได้ของใหม่ทันที ไม่ต้อง build .exe ใหม่
"""
import os
import re
import csv
import json
import asyncio
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
# ลำดับคอลัมน์ในตาราง: Aztek Item Id | ชื่อ | ประเภท | ItemKind | Actions
COL = {'id': 0, 'name': 1, 'type': 2, 'kind': 3}

TMPL_HEADERS = ['ItemKind', 'ItemName', 'duration', 'trade', 'qty']

APP_VERSION = globals().get('TRPU_VERSION') or 'dev'
DATA_DIR = globals().get('TRPU_DATA_DIR') or os.path.join(
    os.environ.get('LOCALAPPDATA') or os.path.expanduser('~'), 'TRPlusUltra')
os.makedirs(DATA_DIR, exist_ok=True)
CHROME_PROFILE = os.path.join(DATA_DIR, 'chrome_profile')
PREF_FILE = os.path.join(DATA_DIR, 'prefs.json')

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
#  [3] Template Excel
# ============================================================================
def build_template(path):
    notes = [
        'ItemKind = รหัสไอเท็ม (ช่องค้นหาหลัก)   |   ItemName = ใส่เพื่อกรองชื่อซ้ำอีกชั้น (ไม่บังคับ)',
        'duration : เว้นว่าง = เฉพาะไอเท็มถาวร  |  ใส่ตัวเลข = จำนวนวันนั้น (เช่น 15)  |  any = ไม่กรอง',
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
    rows = []

    def mk(d):
        return {
            'kind': str(d.get('ItemKind', '') or '').strip(),
            'name': str(d.get('ItemName', '') or '').strip(),
            'dur': str(d.get('duration', '') or '').strip(),
            'trade': norm_bool(d.get('trade', '')),
            'qty': str(d.get('qty', '') or '').strip(),
        }

    if path.lower().endswith(('.xlsx', '.xlsm')):
        if not XLSX_OK:
            raise RuntimeError('เครื่องนี้อ่านไฟล์ Excel ไม่ได้ ลองใช้ .csv แทน')
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        header_row, headers = None, []
        for r in ws.iter_rows(min_row=1, max_row=30):
            vals = [str(c.value or '').strip() for c in r]
            if 'ItemKind' in vals:
                header_row, headers = r[0].row, vals
                break
        if header_row is None:
            wb.close()
            raise RuntimeError('ไม่พบหัวคอลัมน์ ItemKind ในไฟล์')
        idx = {h: i for i, h in enumerate(headers) if h}
        for raw in ws.iter_rows(min_row=header_row + 1, values_only=True):
            if not any(raw):
                continue
            d = {h: (raw[i] if i < len(raw) else '') for h, i in idx.items()}
            row = mk(d)
            if row['kind'] or row['name']:
                rows.append(row)
        wb.close()
        return rows

    with open(path, 'r', encoding='utf-8-sig', errors='replace') as f:
        for d in csv.DictReader(f):
            row = mk(d)
            if row['kind'] or row['name']:
                rows.append(row)
    return rows


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
        self.running = False
        self.cancel = False

        self._build_ui()

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
        self.tab_result = tk.Frame(self.nb, bg=C['bg'])
        self.tab_log = tk.Frame(self.nb, bg=C['bg'])
        self.nb.add(self.tab_search, text='🔍  ค้นหา')
        self.nb.add(self.tab_result, text='📋  ผลลัพธ์')
        self.nb.add(self.tab_log, text='📜  Log')

        self._build_search()
        self._build_result()
        self._build_log()

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

        tk.Label(s3, text='ระยะเวลา: any = ไม่กรอง · เว้นว่าง = เฉพาะไอเทมถาวร · ตัวเลข = จำนวนวันนั้น\n'
                          'ถ้ากรองแค่ระยะเวลาเป็นตัวเลข โปรแกรมจะใช้ฟิลเตอร์บนหน้า list แทน เร็วกว่ามาก',
                 bg=C['bg'], fg=C['dim'], font=('Segoe UI', 8), justify='left'
                 ).grid(row=3, column=0, columnspan=4, sticky='w', pady=(8, 0))

        s4 = tk.Frame(p, bg=C['bg'])
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
        if any(r['id'] == row['id'] for r in self.results):
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

    def pick_file(self):
        path = filedialog.askopenfilename(filetypes=[('Excel / CSV', '*.xlsx *.xlsm *.csv')])
        if not path:
            return
        try:
            self.imported = read_template(path)
            self.lbl_file.config(text=f'{os.path.basename(path)} — {len(self.imported)} รายการ')
            self.btn_multi.config(state='normal' if self.imported else 'disabled')
            self.log(f'อ่าน template: {len(self.imported)} รายการ', 'OK')
        except Exception as ex:
            messagebox.showerror('อ่านไฟล์ไม่ได้', str(ex))

    def save_now(self):
        self.prefs.update({
            'deep': self.v_deep.get(), 'dur': self.v_dur.get().strip(),
            'trade': self.v_trade.get(), 'qty': self.v_qty.get().strip(),
            'headless': self.v_headless.get(),
        })
        save_prefs(self.prefs)

    def deep_criteria(self):
        if not self.v_deep.get():
            return {'dur': 'any', 'trade': 'any', 'qty': ''}
        return {'dur': self.v_dur.get().strip(), 'trade': self.v_trade.get(),
                'qty': self.v_qty.get().strip()}

    def do_cancel(self):
        self.cancel = True
        self.log('กำลังยกเลิก...', 'WARN')

    # ---------- login ----------
    def open_login(self):
        if self.running:
            return messagebox.showinfo('กำลังทำงาน', 'รอให้งานปัจจุบันเสร็จก่อนนะ')
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
        threading.Thread(target=self._thread, args=(criteria,), daemon=True).start()

    def _thread(self, criteria):
        try:
            asyncio.run(self._work(criteria))
        except Exception as ex:
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
            label = f"#{i + 1} Kind={c.get('kind') or '-'}" + (f" Name={c['name']}" if c.get('name') else '')
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

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    App().run()

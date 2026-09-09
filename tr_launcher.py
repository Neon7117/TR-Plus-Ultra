# -*- coding: utf-8 -*-
"""
TR Plus Ultra — Launcher / Auto-updater
=======================================
ตัวเปิดหน้าตาแบบ launcher เกม
  - หลอดโหลดแพทช์ + สถานะการทำงาน
  - แผง PATCH NOTE บอกว่าอัปเดตอะไรมาบ้าง
  - ปุ่ม START เข้าโปรแกรม

*** import playwright/openpyxl ไว้ตรงนี้ด้วย ***
PyInstaller เก็บเฉพาะไลบรารีที่เห็นว่ามีการ import
ถ้าไม่ใส่ โค้ดที่โหลดมาทีหลังจะ import ไม่เจอ
"""
import os
import sys
import json
import time
import threading
import traceback
import urllib.request
import tkinter as tk

# --- บังคับให้ PyInstaller เก็บไลบรารีเหล่านี้เข้าไปใน .exe ---
try:
    import playwright                     # noqa: F401
    from playwright.async_api import async_playwright  # noqa: F401
except Exception:
    pass
try:
    import openpyxl                       # noqa: F401
    from openpyxl.styles import Font, PatternFill, Alignment  # noqa: F401
    from openpyxl.worksheet.datavalidation import DataValidation  # noqa: F401
    from openpyxl.utils import get_column_letter  # noqa: F401
except Exception:
    pass
import asyncio, csv, re, shutil, tkinter.ttk, tkinter.scrolledtext, tkinter.messagebox, tkinter.filedialog  # noqa: F401,E401


# ============================================================================
#  ตั้งค่า — แก้ 2 บรรทัดนี้ให้ชี้ repo ของคุณ
# ============================================================================
GITHUB_USER = 'Neon7117'
GITHUB_REPO = 'TR-Plus-Ultra'
BRANCH = 'main'

MANIFEST_URL = f'https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{BRANCH}/version.json'

LAUNCHER_VERSION = '1.1.0'
APP_NAME = 'TR PLUS ULTRA'
APP_SUB = 'Item Finder for HoF Admin'

DATA_DIR = os.path.join(os.environ.get('LOCALAPPDATA') or os.path.expanduser('~'), 'TRPlusUltra')
os.makedirs(DATA_DIR, exist_ok=True)
APP_FILE = os.path.join(DATA_DIR, 'tr_app.py')
STATE_FILE = os.path.join(DATA_DIR, 'state.json')

# ---- ธีมสี ----
C = {
    'bg':      '#0B1220',
    'panel':   '#131C2E',
    'panel2':  '#0E1626',
    'line':    '#25344F',
    'fg':      '#E6EDF8',
    'dim':     '#8B9BB4',
    'accent':  '#4F8BF0',
    'cyan':    '#35D0BA',
    'green':   '#22C55E',
    'green_d': '#15803D',
    'gold':    '#F5B841',
    'red':     '#EF6461',
    'track':   '#1B2740',
}


# ============================================================================
#  helper
# ============================================================================
def load_state():
    try:
        with open(STATE_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def save_state(d):
    try:
        with open(STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def fetch(url, timeout=15):
    req = urllib.request.Request(
        url + ('&' if '?' in url else '?') + 'ts=' + str(int(time.time())),
        headers={'User-Agent': 'TRPlusUltra-Launcher', 'Cache-Control': 'no-cache'})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode('utf-8')


def vtuple(v):
    out = []
    for part in str(v).split('.'):
        try:
            out.append(int(re.sub(r'\D', '', part) or 0))
        except Exception:
            out.append(0)
    while len(out) < 3:
        out.append(0)
    return tuple(out[:3])


# ============================================================================
#  หน้าต่าง Launcher
# ============================================================================
class Launcher:
    W, H = 920, 552

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(f'{APP_NAME} Launcher')
        self.root.configure(bg=C['bg'])
        self.root.resizable(False, False)
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f'{self.W}x{self.H}+{(sw - self.W) // 2}+{(sh - self.H) // 2 - 20}')

        self.ready = False
        self.launch = False
        self.local_version = load_state().get('version')
        self.new_version = self.local_version
        self.updated_entries = []
        self._pct = 0.0
        self._target = 0.0

        self._build()
        self.root.protocol('WM_DELETE_WINDOW', self._on_close)

    # ---------------- UI ----------------
    def _build(self):
        # ---- แถบหัว ----
        head = tk.Frame(self.root, bg=C['panel2'], height=92)
        head.pack(fill='x')
        head.pack_propagate(False)

        left = tk.Frame(head, bg=C['panel2'])
        left.pack(side='left', padx=26, pady=16)
        tk.Label(left, text=APP_NAME, bg=C['panel2'], fg=C['fg'],
                 font=('Segoe UI', 22, 'bold')).pack(anchor='w')
        tk.Label(left, text=APP_SUB, bg=C['panel2'], fg=C['dim'],
                 font=('Segoe UI', 9)).pack(anchor='w')

        right = tk.Frame(head, bg=C['panel2'])
        right.pack(side='right', padx=26)
        self.badge = tk.Label(right, text='  ●  CONNECTING…  ', bg=C['track'], fg=C['gold'],
                              font=('Segoe UI', 9, 'bold'), padx=6, pady=4)
        self.badge.pack(anchor='e', pady=(20, 4))
        self.ver_lbl = tk.Label(right, text='version —', bg=C['panel2'], fg=C['dim'],
                                font=('Consolas', 9))
        self.ver_lbl.pack(anchor='e')

        tk.Frame(self.root, bg=C['line'], height=1).pack(fill='x')

        # ---- แถบล่าง : หลอดโหลด + START ----
        foot = tk.Frame(self.root, bg=C['panel2'], height=104)
        foot.pack(fill='x', side='bottom')
        foot.pack_propagate(False)
        tk.Frame(self.root, bg=C['line'], height=1).pack(fill='x', side='bottom')

        inner = tk.Frame(foot, bg=C['panel2'])
        inner.pack(fill='both', expand=True, padx=26, pady=18)

        barcol = tk.Frame(inner, bg=C['panel2'])
        barcol.pack(side='left', fill='both', expand=True)

        self.status = tk.Label(barcol, text='กำลังตรวจสอบเวอร์ชัน…', bg=C['panel2'],
                               fg=C['fg'], font=('Segoe UI', 10), anchor='w')
        self.status.pack(fill='x')

        self.canvas = tk.Canvas(barcol, height=16, bg=C['panel2'], bd=0,
                                highlightthickness=0)
        self.canvas.pack(fill='x', pady=(8, 0))
        self.canvas.bind('<Configure>', lambda e: self._draw_bar())

        self.pct_lbl = tk.Label(barcol, text='0%', bg=C['panel2'], fg=C['dim'],
                                font=('Consolas', 9), anchor='w')
        self.pct_lbl.pack(fill='x', pady=(5, 0))

        self.start_btn = tk.Button(inner, text='START  ▶', font=('Segoe UI', 15, 'bold'),
                                   bg=C['track'], fg=C['dim'], bd=0, relief='flat',
                                   activebackground=C['green'], activeforeground='white',
                                   cursor='arrow', state='disabled', command=self._go)
        self.start_btn.pack(side='right', padx=(24, 0), ipadx=30, ipady=13)

        # ---- ตัวกลาง ----
        body = tk.Frame(self.root, bg=C['bg'])
        body.pack(fill='both', expand=True, padx=24, pady=18)

        # ซ้าย : สถานะ
        side = tk.Frame(body, bg=C['bg'], width=286)
        side.pack(side='left', fill='y')
        side.pack_propagate(False)

        box = tk.Frame(side, bg=C['panel'], highlightthickness=1,
                       highlightbackground=C['line'])
        box.pack(fill='both', expand=True)

        tk.Label(box, text='สถานะ', bg=C['panel'], fg=C['dim'],
                 font=('Segoe UI', 9, 'bold')).pack(anchor='w', padx=18, pady=(16, 10))

        self.rows = {}
        for key, text in (('net', 'เชื่อมต่อ GitHub'),
                          ('manifest', 'อ่านข้อมูลเวอร์ชัน'),
                          ('download', 'ดาวน์โหลดโปรแกรม'),
                          ('verify', 'ตรวจไฟล์'),
                          ('done', 'พร้อมใช้งาน')):
            r = tk.Frame(box, bg=C['panel'])
            r.pack(fill='x', padx=18, pady=3)
            dot = tk.Label(r, text='○', bg=C['panel'], fg=C['dim'], font=('Segoe UI', 11))
            dot.pack(side='left')
            lb = tk.Label(r, text=text, bg=C['panel'], fg=C['dim'], font=('Segoe UI', 10))
            lb.pack(side='left', padx=(8, 0))
            self.rows[key] = (dot, lb)

        tk.Frame(box, bg=C['line'], height=1).pack(fill='x', padx=18, pady=(14, 12))
        self.note = tk.Label(box, text='กำลังเริ่มต้น…', bg=C['panel'], fg=C['dim'],
                             font=('Segoe UI', 9), wraplength=230, justify='left')
        self.note.pack(anchor='w', padx=18)

        # ขวา : PATCH NOTE
        rightw = tk.Frame(body, bg=C['bg'])
        rightw.pack(side='left', fill='both', expand=True, padx=(18, 0))

        bar = tk.Frame(rightw, bg=C['accent'], height=34)
        bar.pack(fill='x')
        bar.pack_propagate(False)
        tk.Label(bar, text='PATCH NOTE', bg=C['accent'], fg='white',
                 font=('Segoe UI', 11, 'bold')).pack(side='left', padx=16)
        self.patch_ver = tk.Label(bar, text='', bg=C['accent'], fg='#DCE9FF',
                                  font=('Consolas', 9))
        self.patch_ver.pack(side='right', padx=16)

        wrap = tk.Frame(rightw, bg=C['panel'], highlightthickness=1,
                        highlightbackground=C['line'])
        wrap.pack(fill='both', expand=True)

        self.txt = tk.Text(wrap, bg=C['panel'], fg=C['fg'], bd=0, relief='flat',
                           font=('Segoe UI', 10), wrap='word', padx=18, pady=14,
                           highlightthickness=0, cursor='arrow')
        sb = tk.Scrollbar(wrap, command=self.txt.yview, bg=C['panel'],
                          troughcolor=C['panel2'], bd=0, width=10,
                          activebackground=C['accent'])
        self.txt.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        self.txt.pack(side='left', fill='both', expand=True)

        self.txt.tag_config('ver', foreground=C['cyan'], font=('Segoe UI', 12, 'bold'),
                            spacing1=12, spacing3=2)
        self.txt.tag_config('new', foreground=C['bg'], background=C['gold'],
                            font=('Segoe UI', 8, 'bold'))
        self.txt.tag_config('date', foreground=C['dim'], font=('Consolas', 9))
        self.txt.tag_config('item', lmargin1=16, lmargin2=32, spacing3=4)
        self.txt.tag_config('dim', foreground=C['dim'])
        self.txt.insert('end', '\nกำลังโหลดรายการอัปเดต…', 'dim')
        self.txt.config(state='disabled')

        self._tick_bar()

    # ---------------- วาดหลอดโหลด ----------------
    def _draw_bar(self):
        c = self.canvas
        c.delete('all')
        w = c.winfo_width() or 1
        h = 16
        c.create_rectangle(0, 3, w, h - 1, fill=C['track'], outline='')
        fw = int(w * max(0.0, min(1.0, self._pct / 100.0)))
        if fw > 0:
            c.create_rectangle(0, 3, fw, h - 1, fill=C['cyan'], outline='')
            if fw > 6:
                c.create_rectangle(max(0, fw - 5), 3, fw, h - 1, fill='#8CF0E1', outline='')

    def _tick_bar(self):
        if abs(self._target - self._pct) > 0.4:
            self._pct += (self._target - self._pct) * 0.18
        else:
            self._pct = self._target
        self._draw_bar()
        self.pct_lbl.config(text=f'{int(round(self._pct))}%'
                            + ('   Complete' if self._pct >= 99.5 else ''))
        self.root.after(30, self._tick_bar)

    # ---------------- อัปเดต UI (เรียกจากเธรดอื่นได้) ----------------
    def ui(self, fn):
        try:
            self.root.after(0, fn)
        except Exception:
            pass

    def set_step(self, key, state):
        """state: run / ok / skip / fail"""
        def _do():
            dot, lb = self.rows[key]
            if state == 'run':
                dot.config(text='◐', fg=C['gold']); lb.config(fg=C['fg'])
            elif state == 'ok':
                dot.config(text='●', fg=C['green']); lb.config(fg=C['fg'])
            elif state == 'skip':
                dot.config(text='–', fg=C['dim']); lb.config(fg=C['dim'])
            else:
                dot.config(text='✕', fg=C['red']); lb.config(fg=C['red'])
        self.ui(_do)

    def set_status(self, text, pct=None, note=None):
        def _do():
            self.status.config(text=text)
            if pct is not None:
                self._target = float(pct)
            if note is not None:
                self.note.config(text=note)
        self.ui(_do)

    def set_badge(self, text, color):
        self.ui(lambda: self.badge.config(text=f'  ●  {text}  ', fg=color))

    def set_version(self, v):
        self.ui(lambda: self.ver_lbl.config(text=f'version {v}'))
        self.ui(lambda: self.patch_ver.config(text=f'ล่าสุด {v}'))

    def fill_patch(self, history, newest=None):
        def _do():
            self.txt.config(state='normal')
            self.txt.delete('1.0', 'end')
            if not history:
                self.txt.insert('end', '\nยังไม่มีข้อมูลอัปเดต', 'dim')
            for e in history[:12]:
                v = e.get('version', '?')
                self.txt.insert('end', f'เวอร์ชัน {v}', 'ver')
                if newest and v == newest:
                    self.txt.insert('end', '  ')
                    self.txt.insert('end', ' ใหม่ ', 'new')
                if e.get('date'):
                    self.txt.insert('end', f'   {e["date"]}', 'date')
                self.txt.insert('end', '\n')
                for ch in e.get('changes', []):
                    self.txt.insert('end', f'  •  {ch}\n', 'item')
            self.txt.config(state='disabled')
            self.txt.yview_moveto(0)
        self.ui(_do)

    def enable_start(self, label='START  ▶'):
        def _do():
            self.ready = True
            self.start_btn.config(state='normal', bg=C['green'], fg='white',
                                  cursor='hand2', text=label,
                                  activebackground=C['green_d'])
            self.start_btn.focus_set()
        self.ui(_do)
        self.root.bind('<Return>', lambda e: self._go())

    # ---------------- ปุ่ม ----------------
    def _go(self):
        if not self.ready:
            return
        self.launch = True
        self.root.destroy()

    def _on_close(self):
        self.launch = False
        self.root.destroy()

    # ---------------- งานเบื้องหลัง ----------------
    def work(self):
        state = load_state()
        local = state.get('version')
        self.set_version(local or '—')

        self.set_step('net', 'run')
        self.set_status('กำลังเชื่อมต่อ GitHub…', 12, 'ติดต่อเซิร์ฟเวอร์อัปเดต')
        manifest, err = None, None
        try:
            raw = fetch(MANIFEST_URL)
            manifest = json.loads(raw)
        except Exception as ex:
            err = str(ex)

        if not manifest:
            self.set_step('net', 'fail')
            self.set_badge('OFFLINE', C['red'])
            if os.path.exists(APP_FILE):
                self.set_step('manifest', 'skip')
                self.set_step('download', 'skip')
                self.set_step('verify', 'ok')
                self.set_step('done', 'ok')
                self.set_status('ออฟไลน์ — ใช้เวอร์ชันที่โหลดไว้ในเครื่อง', 100,
                                'ต่อเน็ตไม่ได้ จึงใช้ตัวที่โหลดไว้ล่าสุด ทำงานได้ตามปกติ')
                self.fill_patch([{'version': local or '?', 'date': '',
                                  'changes': ['ออฟไลน์อยู่ ยังไม่ได้เช็กว่ามีเวอร์ชันใหม่ไหม']}])
                self.new_version = local
                self.enable_start()
            else:
                self.set_status('เชื่อมต่อไม่ได้ และยังไม่เคยโหลดโปรแกรม', 0,
                                'กรุณาเช็กอินเทอร์เน็ตแล้วเปิดใหม่')
                self.fill_patch([{'version': '—', 'date': '',
                                  'changes': ['ต่อ GitHub ไม่ได้: ' + (err or '')[:120]]}])
            return

        self.set_step('net', 'ok')
        self.set_badge('ONLINE', C['green'])

        self.set_step('manifest', 'run')
        self.set_status('อ่านข้อมูลเวอร์ชัน…', 34)
        remote = str(manifest.get('version', '0.0.0'))
        history = manifest.get('history', [])
        app_url = manifest.get('app_url') or (
            f'https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{BRANCH}/app/tr_app.py')
        self.set_step('manifest', 'ok')
        self.set_version(remote)

        need = (not os.path.exists(APP_FILE)) or (local is None) or (vtuple(remote) > vtuple(local))

        if need:
            self.set_step('download', 'run')
            self.set_status(f'กำลังดาวน์โหลดเวอร์ชัน {remote}…', 62,
                            f'มีเวอร์ชันใหม่ {remote}\nกำลังอัปเดตให้อัตโนมัติ')
            try:
                code = fetch(app_url, timeout=60)
                self.set_step('download', 'ok')
                self.set_step('verify', 'run')
                self.set_status('ตรวจไฟล์…', 86)
                if 'TR_APP_MARKER' not in code:
                    raise ValueError('ไฟล์ที่โหลดมาไม่ใช่โปรแกรมที่ถูกต้อง')
                with open(APP_FILE, 'w', encoding='utf-8') as f:
                    f.write(code)
                self.set_step('verify', 'ok')
                if local is None:
                    self.updated_entries = history[:1]
                else:
                    self.updated_entries = [h for h in history
                                            if vtuple(h.get('version', '0')) > vtuple(local)]
                state['version'] = remote
                save_state(state)
                self.new_version = remote
                self.fill_patch(history, newest=remote)
                self.set_step('done', 'ok')
                self.set_status(f'อัปเดตเป็นเวอร์ชัน {remote} เรียบร้อย', 100,
                                'อัปเดตเสร็จแล้ว ดูรายการที่แก้ได้ทางขวา')
                self.enable_start('START  ▶')
            except Exception as ex:
                self.set_step('download', 'fail')
                if os.path.exists(APP_FILE):
                    self.set_step('verify', 'skip')
                    self.set_step('done', 'ok')
                    self.new_version = local
                    self.fill_patch(history, newest=remote)
                    self.set_status('อัปเดตไม่สำเร็จ — ใช้เวอร์ชันเดิมไปก่อน', 100,
                                    'โหลดไม่สำเร็จ: ' + str(ex)[:90])
                    self.enable_start()
                else:
                    self.set_status('อัปเดตไม่สำเร็จ และยังไม่มีโปรแกรมในเครื่อง', 0,
                                    str(ex)[:140])
                    self.fill_patch(history, newest=remote)
            return

        # ---- ไม่มีอะไรใหม่ ----
        self.set_step('download', 'skip')
        self.set_step('verify', 'ok')
        self.set_step('done', 'ok')
        self.new_version = remote
        self.fill_patch(history, newest=remote)
        self.set_status(f'เวอร์ชัน {remote} — ล่าสุดแล้ว', 100,
                        'ไม่มีอัปเดตใหม่ พร้อมใช้งานได้เลย')
        self.enable_start()

    def run(self):
        threading.Thread(target=self._work_guard, daemon=True).start()
        self.root.mainloop()
        return self.launch

    def _work_guard(self):
        try:
            self.work()
        except Exception:
            self.set_status('เกิดข้อผิดพลาดในตัวเปิด', 0, traceback.format_exc()[-160:])


# ============================================================================
#  หน้าต่างแจ้ง error
# ============================================================================
def error_box(title, message):
    r = tk.Tk()
    r.title(title)
    r.configure(bg=C['bg'])
    w, h = 520, 250
    sw, sh = r.winfo_screenwidth(), r.winfo_screenheight()
    r.geometry(f'{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}')
    tk.Label(r, text=title, bg=C['bg'], fg=C['red'],
             font=('Segoe UI', 14, 'bold')).pack(anchor='w', padx=24, pady=(22, 8))
    tk.Label(r, text=message, bg=C['bg'], fg=C['fg'], font=('Segoe UI', 10),
             wraplength=w - 50, justify='left').pack(anchor='w', padx=24)
    tk.Button(r, text='ปิด', bg=C['panel'], fg=C['fg'], bd=0, font=('Segoe UI', 10),
              cursor='hand2', command=r.destroy).pack(side='bottom', fill='x',
                                                      padx=24, pady=20, ipady=8)
    r.mainloop()


# ============================================================================
#  main
# ============================================================================
def main():
    lc = Launcher()
    launch = lc.run()
    if not launch:
        return

    try:
        with open(APP_FILE, 'r', encoding='utf-8') as f:
            code = f.read()
        g = {
            '__name__': '__main__',
            '__file__': APP_FILE,
            'TRPU_VERSION': lc.new_version or '?',
            'TRPU_DATA_DIR': DATA_DIR,
            'TRPU_LAUNCHER': LAUNCHER_VERSION,
        }
        exec(compile(code, APP_FILE, 'exec'), g)
    except Exception:
        error_box('โปรแกรมมีปัญหา',
                  'เปิดโปรแกรมไม่สำเร็จ กรุณาส่งข้อความนี้ให้คนดูแลเครื่องมือ\n\n'
                  + traceback.format_exc()[-900:])


if __name__ == '__main__':
    main()

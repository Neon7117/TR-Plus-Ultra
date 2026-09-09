# -*- coding: utf-8 -*-
"""
TR Plus Ultra — Launcher / Auto-updater
=======================================
ไฟล์นี้คือ "ตัวเปิด" ที่จะถูก build เป็น .exe แจกทีมครั้งเดียว
หน้าที่ของมัน:
  1. เช็ก version.json บน GitHub ว่ามีเวอร์ชันใหม่ไหม
  2. ถ้ามี → โหลดโค้ดใหม่มา แล้วขึ้นหน้าต่างบอกว่าแก้อะไรมาบ้าง
  3. ถ้าไม่มี → ขึ้นแถบสั้นๆ บอกว่าเป็นเวอร์ชันล่าสุดแล้ว
  4. เปิดโปรแกรมตัวจริง

*** สำคัญ: import พวก playwright/openpyxl ไว้ตรงนี้ด้วย ***
เพราะ PyInstaller จะเก็บเฉพาะไลบรารีที่เห็นว่ามีการ import
ถ้าไม่ใส่ โค้ดที่โหลดมาทีหลังจะ import ไม่เจอ
"""
import os
import sys
import json
import time
import traceback
import urllib.request
import tkinter as tk
from tkinter import ttk

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
except Exception:
    pass
import asyncio, threading, csv, re, shutil  # noqa: F401,E401


# ============================================================================
#  ตั้งค่า — แก้ 2 บรรทัดนี้ให้ชี้ repo ของคุณ
# ============================================================================
GITHUB_USER = 'CHANGE-ME'
GITHUB_REPO = 'tr-plus-ultra'
BRANCH = 'main'

MANIFEST_URL = f'https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{BRANCH}/version.json'

LAUNCHER_VERSION = '1.0.0'
APP_NAME = 'TR Plus Ultra'

# ---- ที่เก็บไฟล์ในเครื่อง ----
DATA_DIR = os.path.join(os.environ.get('LOCALAPPDATA') or os.path.expanduser('~'), 'TRPlusUltra')
os.makedirs(DATA_DIR, exist_ok=True)
APP_FILE = os.path.join(DATA_DIR, 'tr_app.py')
STATE_FILE = os.path.join(DATA_DIR, 'state.json')

# ---- สี (ให้เข้ากับตัวโปรแกรมหลัก) ----
C = {
    'bg': '#11161f', 'card': '#151c28', 'line': '#263041',
    'fg': '#dfe6f2', 'dim': '#8b98ad', 'accent': '#4f6bed',
    'ok': '#5ed09a', 'warn': '#e3b341', 'err': '#f0736a',
}


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


def fetch(url, timeout=10):
    req = urllib.request.Request(
        url + ('&' if '?' in url else '?') + 'ts=' + str(int(time.time())),
        headers={'User-Agent': 'TRPlusUltra-Launcher', 'Cache-Control': 'no-cache'})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode('utf-8')


def vtuple(v):
    """'0.2.10' -> (0,2,10) เอาไว้เทียบว่าใหม่กว่าไหม"""
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
#  หน้าต่าง UI ของตัวเปิด
# ============================================================================
class Splash:
    """แถบเล็กๆ ตอนเปิดโปรแกรม บอกสถานะและเวอร์ชัน"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.configure(bg=C['line'])
        w, h = 380, 108
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f'{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}')

        wrap = tk.Frame(self.root, bg=C['card'])
        wrap.place(x=1, y=1, width=w - 2, height=h - 2)

        tk.Label(wrap, text=APP_NAME, bg=C['card'], fg=C['fg'],
                 font=('Segoe UI', 15, 'bold')).pack(pady=(20, 2))
        self.msg = tk.Label(wrap, text='กำลังตรวจสอบเวอร์ชัน...', bg=C['card'],
                            fg=C['dim'], font=('Segoe UI', 10))
        self.msg.pack()
        self.bar = ttk.Progressbar(wrap, mode='indeterminate', length=250)
        self.bar.pack(pady=(12, 0))
        self.bar.start(14)
        self.root.update()

    def say(self, text, color=None):
        self.msg.config(text=text, fg=color or C['dim'])
        self.root.update()

    def close(self):
        try:
            self.bar.stop()
            self.root.destroy()
        except Exception:
            pass


class ChangelogWindow:
    """หน้าต่างบอกว่าอัปเดตอะไรมาบ้าง — กด 'เริ่มใช้งาน' เพื่อไปต่อ"""

    def __init__(self, new_version, entries):
        self.root = tk.Tk()
        self.root.title(f'{APP_NAME} — อัปเดตแล้ว')
        self.root.configure(bg=C['bg'])
        w, h = 520, 430
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f'{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}')
        self.root.resizable(False, False)

        head = tk.Frame(self.root, bg=C['card'], height=76)
        head.pack(fill='x')
        head.pack_propagate(False)
        tk.Label(head, text='อัปเดตเรียบร้อย', bg=C['card'], fg=C['ok'],
                 font=('Segoe UI', 16, 'bold')).pack(anchor='w', padx=22, pady=(16, 0))
        tk.Label(head, text=f'ตอนนี้เป็นเวอร์ชัน {new_version}', bg=C['card'], fg=C['dim'],
                 font=('Segoe UI', 10)).pack(anchor='w', padx=22)

        body = tk.Frame(self.root, bg=C['bg'])
        body.pack(fill='both', expand=True, padx=22, pady=(16, 0))

        tk.Label(body, text='มีอะไรใหม่', bg=C['bg'], fg=C['dim'],
                 font=('Segoe UI', 9, 'bold')).pack(anchor='w')

        canvas = tk.Text(body, bg=C['card'], fg=C['fg'], bd=0, relief='flat',
                         font=('Segoe UI', 10), wrap='word', padx=14, pady=12,
                         highlightthickness=1, highlightbackground=C['line'])
        canvas.pack(fill='both', expand=True, pady=(6, 0))
        canvas.tag_config('ver', foreground=C['accent'], font=('Segoe UI', 11, 'bold'),
                          spacing1=10, spacing3=4)
        canvas.tag_config('date', foreground=C['dim'], font=('Segoe UI', 9))
        canvas.tag_config('item', lmargin1=14, lmargin2=26, spacing3=3)

        for e in entries:
            canvas.insert('end', f"เวอร์ชัน {e.get('version', '?')}", 'ver')
            if e.get('date'):
                canvas.insert('end', f"   {e['date']}\n", 'date')
            else:
                canvas.insert('end', '\n')
            for ch in e.get('changes', []):
                canvas.insert('end', f'  •  {ch}\n', 'item')
        canvas.config(state='disabled')

        foot = tk.Frame(self.root, bg=C['bg'])
        foot.pack(fill='x', padx=22, pady=16)
        btn = tk.Button(foot, text='เริ่มใช้งาน', bg=C['accent'], fg='white', bd=0,
                        font=('Segoe UI', 11, 'bold'), cursor='hand2',
                        activebackground='#3d55cf', activeforeground='white',
                        command=self.root.destroy)
        btn.pack(fill='x', ipady=9)
        self.root.bind('<Return>', lambda e: self.root.destroy())
        btn.focus_set()

    def show(self):
        self.root.mainloop()


def error_box(title, message):
    r = tk.Tk()
    r.title(title)
    r.configure(bg=C['bg'])
    w, h = 470, 210
    sw, sh = r.winfo_screenwidth(), r.winfo_screenheight()
    r.geometry(f'{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}')
    tk.Label(r, text=title, bg=C['bg'], fg=C['err'],
             font=('Segoe UI', 13, 'bold')).pack(anchor='w', padx=22, pady=(20, 6))
    tk.Label(r, text=message, bg=C['bg'], fg=C['fg'], font=('Segoe UI', 10),
             wraplength=w - 46, justify='left').pack(anchor='w', padx=22)
    tk.Button(r, text='ปิด', bg=C['card'], fg=C['fg'], bd=0, font=('Segoe UI', 10),
              cursor='hand2', command=r.destroy).pack(side='bottom', fill='x',
                                                      padx=22, pady=18, ipady=7)
    r.mainloop()


# ============================================================================
#  ตัวหลัก
# ============================================================================
def main():
    splash = Splash()
    state = load_state()
    local_version = state.get('version')
    manifest = None
    offline_reason = None

    try:
        manifest = json.loads(fetch(MANIFEST_URL))
    except Exception as ex:
        offline_reason = str(ex)

    updated_entries = None
    new_version = local_version

    if manifest:
        remote_version = str(manifest.get('version', '0.0.0'))
        app_url = manifest.get('app_url') or (
            f'https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{BRANCH}/app/tr_app.py')

        need = (not os.path.exists(APP_FILE)) or (local_version is None) \
            or (vtuple(remote_version) > vtuple(local_version))

        if need:
            splash.say(f'กำลังดาวน์โหลดเวอร์ชัน {remote_version} ...', C['accent'])
            try:
                code = fetch(app_url, timeout=30)
                if 'TR_APP_MARKER' not in code:
                    raise ValueError('ไฟล์ที่โหลดมาไม่ใช่โปรแกรมที่ถูกต้อง')
                with open(APP_FILE, 'w', encoding='utf-8') as f:
                    f.write(code)
                # เก็บเฉพาะรายการที่ใหม่กว่าเวอร์ชันเดิม เอาไว้โชว์
                hist = manifest.get('history', [])
                if local_version is None:
                    updated_entries = hist[:1]
                else:
                    updated_entries = [h for h in hist
                                       if vtuple(h.get('version', '0')) > vtuple(local_version)]
                if not updated_entries:
                    updated_entries = hist[:1]
                state['version'] = remote_version
                save_state(state)
                new_version = remote_version
            except Exception as ex:
                offline_reason = str(ex)
                if not os.path.exists(APP_FILE):
                    splash.close()
                    error_box('อัปเดตไม่สำเร็จ',
                              'โหลดโปรแกรมจาก GitHub ไม่ได้ และยังไม่มีตัวที่เคยโหลดไว้ในเครื่อง\n\n'
                              f'รายละเอียด: {ex}')
                    return
        else:
            splash.say(f'เวอร์ชัน {remote_version} — ล่าสุดแล้ว', C['ok'])
            new_version = remote_version
            time.sleep(1.1)
    else:
        if not os.path.exists(APP_FILE):
            splash.close()
            error_box('เชื่อมต่อไม่ได้',
                      'ต่อ GitHub ไม่ได้ และยังไม่เคยโหลดโปรแกรมมาก่อน\n'
                      'กรุณาเช็กอินเทอร์เน็ตแล้วเปิดใหม่อีกครั้ง\n\n'
                      f'รายละเอียด: {offline_reason}')
            return
        splash.say(f'ออฟไลน์ — ใช้เวอร์ชัน {local_version} ที่โหลดไว้', C['warn'])
        time.sleep(1.4)

    splash.close()

    if updated_entries:
        ChangelogWindow(new_version, updated_entries).show()

    # ---- รันโปรแกรมตัวจริง ----
    try:
        with open(APP_FILE, 'r', encoding='utf-8') as f:
            code = f.read()
        g = {
            '__name__': '__main__',
            '__file__': APP_FILE,
            'TRPU_VERSION': new_version,
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

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
svchosts.py
USB kontrol + kilit ekranı uygulaması
- Supabase entegrasyonu
- Alt+F4 engelleme
- Kullanıcı onayıyla admin yükseltme
- İlk çalıştırmada kendini Program Files içine TAŞIMA
- allowed_serials.json daima Program Files içinde tutulur
- Başlangıca kısayol ekleme
"""

import os
import sys
import json
import time
import threading
import subprocess
import ctypes
import psutil
import signal
import tkinter as tk
import tkinter.messagebox as messagebox
import shutil
import requests

# ==============
# Opsiyonel libler
# ==============
try:
    from plyer import notification as plyer_notification
    _HAS_PLYER = True
except Exception:
    _HAS_PLYER = False

try:
    from win10toast import ToastNotifier
    _HAS_WIN10TOAST = True
except Exception:
    _HAS_WIN10TOAST = False

try:
    import keyboard
    _HAS_KEYBOARD = True
except Exception:
    _HAS_KEYBOARD = False

try:
    import pythoncom
    from win32com.client import Dispatch
    _HAS_PYWIN32 = True
except Exception:
    _HAS_PYWIN32 = False

try:
    import winreg
    _HAS_WINREG = True
except Exception:
    winreg = None
    _HAS_WINREG = False

# =================
# Config
# =================
PROG_DIR_NAME = "Svchosts"
PF_DIR = os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), PROG_DIR_NAME)
os.makedirs(PF_DIR, exist_ok=True)

ALLOWED_SERIALS_FILE = os.path.join(PF_DIR, "allowed_serials.json")
HOTKEY_PASSWORD = "yuzuk123."
NOTIFY_DURATION = 5
POLL_INTERVAL = 2
SHUTDOWN_INTERVAL = 160

# Supabase
DB_URL = "https://ubncrpchxqtgybnwyasi.supabase.co/rest/v1/data_table"
SERVICE_KEY = (
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InVibmNycGNoeHF0Z3libnd5YXNpIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc1ODkxMDU1NCwiZXhwIjoyMDc0NDg2NTU0fQ.XwH8JbU3JiQeJDNmM13sJFQQ_Sen8e005VeqM6RdUEM"
)

# =================
# Globals
# =================
_allowed_lock = threading.Lock()
allowed_serials = []
_lockscreen_shown = False
_lockscreen_obj = None
_shutdown_triggered = False

# =================
# Helpers
# =================
def current_executable_path():
    if getattr(sys, "frozen", False):
        return os.path.abspath(sys.executable)
    return os.path.abspath(sys.argv[0])

def is_running_as_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False

def fetch_database():
    headers = {
        "apikey": SERVICE_KEY,
        "Authorization": f"Bearer {SERVICE_KEY}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    try:
        resp = requests.get(f"{DB_URL}?select=*", headers=headers, timeout=10)
        resp.raise_for_status()
        rows = resp.json()
        if isinstance(rows, list):
            return [str(r.get("value", "")).upper() for r in rows if "value" in r]
    except Exception as e:
        print(f"[WARN] fetch_database: {e}")
    return []

# Notification
_toast = None
if _HAS_WIN10TOAST:
    try:
        _toast = ToastNotifier()
    except Exception:
        _toast = None

def show_notification(title, msg, duration=NOTIFY_DURATION):
    try:
        if _HAS_PLYER:
            plyer_notification.notify(title=title, message=msg, timeout=duration)
        elif _toast:
            _toast.show_toast(title, msg, duration=duration, threaded=True)
        else:
            print(f"[NOTIFY] {title}: {msg}")
    except Exception:
        pass

# Serial file
def load_allowed_serials():
    global allowed_serials
    try:
        if os.path.exists(ALLOWED_SERIALS_FILE):
            with open(ALLOWED_SERIALS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    with _allowed_lock:
                        allowed_serials = [str(x).upper() for x in data]
                    return
    except Exception as e:
        print(f"[WARN] load_allowed_serials: {e}")
    with _allowed_lock:
        allowed_serials = []

def save_allowed_serials():
    try:
        with _allowed_lock:
            data = list(allowed_serials)
        with open(ALLOWED_SERIALS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        print(f"[WARN] save_allowed_serials: {e}")

# USB detect
def _run_cmd(cmd):
    try:
        return subprocess.check_output(cmd, shell=True).decode(errors="ignore")
    except Exception:
        return ""

def get_usb_serials():
    serials = set()
    out = _run_cmd('wmic path Win32_DiskDrive where "InterfaceType=\'USB\'" get PNPDeviceID')
    for line in out.splitlines():
        line = line.strip()
        if not line or "PNPDeviceID" in line: continue
        if "USBSTOR" not in line.upper(): continue
        try:
            part = line.split("\\")[-1].strip()
            if not part: continue
            serial_raw = part.split("&")[0]
            cleaned = "".join(ch for ch in serial_raw if ch.isalnum()).upper()
            if cleaned: serials.add(cleaned)
        except: continue
    return list(serials)

def shutdown_pc():
    global _shutdown_triggered
    if _shutdown_triggered: return
    _shutdown_triggered = True
    os.system("shutdown /s /t 0")

# Taskbar
def hide_taskbar():
    hwnd = ctypes.windll.user32.FindWindowW("Shell_TrayWnd", None)
    if hwnd: ctypes.windll.user32.ShowWindow(hwnd, 0)

def show_taskbar():
    hwnd = ctypes.windll.user32.FindWindowW("Shell_TrayWnd", None)
    if hwnd: ctypes.windll.user32.ShowWindow(hwnd, 5)

# Lock screen
def animate_rgb(lbl, step=0):
    import colorsys
    hue = (step % 360) / 360.0
    r,g,b = colorsys.hsv_to_rgb(hue,1,1)
    lbl.config(fg=f"#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}")
    lbl.after(60, lambda: animate_rgb(lbl, step+3))

class LockScreen:
    def __init__(self, countdown): self.countdown=int(countdown); self._running=False; self._window=None
    def show(self):
        if self._running: return
        self._running=True
        threading.Thread(target=self._run, daemon=True).start()
    def hide(self):
        self._running=False
        if self._window: self._window.after(0, self._window.destroy); self._window=None
    def _run(self):
        root=tk.Tk(); self._window=root
        root.title("Ekran Kilidi"); root.attributes("-topmost",True)
        try: root.attributes("-fullscreen",True)
        except: root.geometry(f"{root.winfo_screenwidth()}x{root.winfo_screenheight()}+0+0")
        root.config(cursor="none", bg="#1a1a1a")
        root.protocol("WM_DELETE_WINDOW", lambda: None)
        root.bind("<Alt-F4>", lambda e:"break")
        f=tk.Frame(root,bg="#1a1a1a"); f.pack(expand=True,fill="both")
        tk.Label(f,text="EKRAN KİLİTLİ",font=("Segoe UI",48,"bold"),fg="white",bg="#1a1a1a").pack(pady=20)
        msg=tk.Label(f,text="USB takınız",font=("Segoe UI",24),fg="white",bg="#1a1a1a"); msg.pack(pady=10)
        cnt=tk.Label(f,text=f"Kapatma: {self.countdown}s",font=("Segoe UI",36),fg="white",bg="#1a1a1a"); cnt.pack(pady=10)
        footer=tk.Label(f,text="Yapımcılar: Burak & Hüseyin",font=("Segoe UI",14),bg="#1a1a1a"); footer.pack(side="bottom",pady=20); animate_rgb(footer)
        def tick():
            if not self._running: root.destroy(); return
            cnt.config(text=f"Kapatma: {self.countdown}s"); self.countdown-=1
            if self.countdown<=0: shutdown_pc()
            root.after(1000,tick)
        root.after(0,tick); root.mainloop()

# Startup shortcut
def create_startup_shortcut_for_target(target):
    startup_dir=os.path.join(os.environ.get("APPDATA",""),r"Microsoft\Windows\Start Menu\Programs\Startup")
    os.makedirs(startup_dir,exist_ok=True)
    link=os.path.join(startup_dir,os.path.splitext(os.path.basename(target))[0]+".lnk")
    if _HAS_PYWIN32:
        pythoncom.CoInitialize()
        shell=Dispatch('WScript.Shell'); sc=shell.CreateShortcut(link)
        sc.TargetPath=target; sc.WorkingDirectory=os.path.dirname(target); sc.save()
    else:
        with open(link+".bat","w") as f: f.write(f'start "" "{target}"\n')

# Install / move
def perform_elevated_install():
    try:
        src=current_executable_path()
        dest_dir=PF_DIR; os.makedirs(dest_dir,exist_ok=True)
        dest_path=os.path.join(dest_dir,os.path.basename(src))
        if os.path.abspath(os.path.dirname(src)).lower()==os.path.abspath(dest_dir).lower():
            return True
        if os.path.exists(dest_path): os.remove(dest_path)
        shutil.move(src,dest_path)
        # migrate serials.json if exists on desktop
        desktop=os.path.join(os.path.expanduser("~"),"Desktop","allowed_serials.json")
        if os.path.exists(desktop): shutil.move(desktop,ALLOWED_SERIALS_FILE)
        create_startup_shortcut_for_target(dest_path)
        os.execv(dest_path,[dest_path])
        return True
    except Exception as e:
        print(f"[WARN] install: {e}"); return False

def prompt_install_with_elevation():
    root=tk.Tk(); root.withdraw()
    if not messagebox.askyesno("Kurulum","Program Files içine taşınıp Başlangıca eklenmesini ister misiniz?"):
        root.destroy(); return False
    root.destroy()
    if is_running_as_admin():
        return perform_elevated_install()
    try:
        ctypes.windll.shell32.ShellExecuteW(None,"runas",sys.executable,f'"{current_executable_path()}" --elevated-install',None,1)
        return True
    except Exception as e:
        print(f"[WARN] elevation: {e}"); return False

# Monitor
def monitor_loop():
    global _lockscreen_shown,_lockscreen_obj
    while True:
        time.sleep(POLL_INTERVAL)
        serials=get_usb_serials()
        with _allowed_lock:
            valid=any(s in allowed_serials for s in serials)
        if valid:
            if _lockscreen_shown and _lockscreen_obj: _lockscreen_obj.hide(); _lockscreen_shown=False
            show_taskbar()
        else:
            if not _lockscreen_shown:
                _lockscreen_obj=LockScreen(SHUTDOWN_INTERVAL); _lockscreen_obj.show(); _lockscreen_shown=True
            hide_taskbar()

# Main
def main():
    if "--elevated-install" in sys.argv:
        perform_elevated_install(); return
    src_dir=os.path.abspath(os.path.dirname(current_executable_path()))
    if not src_dir.lower().startswith(PF_DIR.lower()):
        prompt_install_with_elevation(); return
    load_allowed_serials()
    db_values=fetch_database()
    if db_values:
        with _allowed_lock:
            for v in db_values:
                if v and v not in allowed_serials: allowed_serials.append(v)
        save_allowed_serials()
    threading.Thread(target=monitor_loop,daemon=True).start()
    show_notification("USB Kontrol","Servis başlatıldı")
    while True: time.sleep(1)

if __name__=="__main__":
    main()

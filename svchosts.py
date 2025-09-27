#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
svchosts.py
USB kontrol + kilit ekranı uygulaması (Supabase + Alt+F4 engelleme + user-consent install).
Çalıştırma: python svchosts.py
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
from datetime import datetime
import requests   # for Supabase fetch

# optional libs
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

# pywin32 for shortcut creation (optional)
try:
    import pythoncom
    from win32com.client import Dispatch
    _HAS_PYWIN32 = True
except Exception:
    _HAS_PYWIN32 = False

# Registry (Windows)
try:
    import winreg
    _HAS_WINREG = True
except Exception:
    winreg = None
    _HAS_WINREG = False

# =========================
# CONFIG
# =========================
PROG_DIR_NAME = "Svchosts"
EXE_NAME = "svchosts.exe"  # used if you are packaging as exe
POLL_INTERVAL = 2          # seconds
SHUTDOWN_INTERVAL = 160    # seconds before shutdown when USB missing
ALLOWED_SERIALS_FILE = "allowed_serials.json"
HOTKEY_PASSWORD = "yuzuk123."
NOTIFY_DURATION = 5

# Supabase remote table config (unchanged)
DB_URL = "https://ubncrpchxqtgybnwyasi.supabase.co/rest/v1/data_table"
SERVICE_KEY = (
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InVibmNycGNoeHF0Z3libnd5YXNpIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc1ODkxMDU1NCwiZXhwIjoyMDc0NDg2NTU0fQ.XwH8JbU3JiQeJDNmM13sJFQQ_Sen8e005VeqM6RdUEM"
)

def fetch_database():
    """
    Fetch rows from Supabase REST endpoint and return a list of uppercased 'value' strings.
    Returns [] on any error.
    """
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
            values = []
            for row in rows:
                if isinstance(row, dict) and "value" in row:
                    v = row.get("value") or ""
                    v = str(v).upper()
                    if v:
                        values.append(v)
            return values
    except Exception as e:
        print(f"[WARN] fetch_database error: {e}")
    return []

# Globals
_allowed_lock = threading.Lock()
allowed_serials = []
_lockscreen_shown = False
_lockscreen_obj = None
_shutdown_triggered = False

# Notification helper
_toast = None
if _HAS_WIN10TOAST:
    try:
        _toast = ToastNotifier()
    except Exception:
        _toast = None

def show_notification(title, message, duration=NOTIFY_DURATION):
    try:
        if _HAS_PLYER:
            plyer_notification.notify(title=title, message=message, timeout=duration)
            return
        if _toast:
            _toast.show_toast(title, message, duration=duration, threaded=True)
            return
        print(f"[NOTIF] {title}: {message}")
    except Exception:
        pass

# allowed_serials persistence
def load_allowed_serials():
    global allowed_serials
    try:
        if os.path.exists(ALLOWED_SERIALS_FILE):
            with open(ALLOWED_SERIALS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    with _allowed_lock:
                        allowed_serials = [str(x).upper() for x in data]
                    print(f"[INFO] {ALLOWED_SERIALS_FILE} loaded: {allowed_serials}")
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
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(f"[INFO] saved {ALLOWED_SERIALS_FILE}")
    except Exception as e:
        print(f"[WARN] save_allowed_serials: {e}")

# USB serial detection (USBSTOR)
def _run_cmd(cmd):
    try:
        out = subprocess.check_output(cmd, shell=True)
        return out.decode(errors="ignore")
    except Exception:
        return ""

def get_usb_serials():
    serials = set()
    try:
        out = _run_cmd('wmic path Win32_DiskDrive where "InterfaceType=\'USB\'" get PNPDeviceID')
        for line in out.splitlines():
            line = line.strip()
            if not line or "PNPDeviceID" in line:
                continue
            if "USBSTOR" not in line.upper():
                continue
            try:
                part = line.split("\\")[-1].strip()
                if not part:
                    continue
                serial_raw = part.split("&")[0]
                cleaned = "".join(ch for ch in serial_raw if ch.isalnum()).upper()
                if cleaned:
                    serials.add(cleaned)
            except Exception:
                continue
    except Exception as e:
        print(f"[WARN] get_usb_serials: {e}")
    return list(serials)

# Process / shutdown
def shutdown_pc():
    global _shutdown_triggered
    if _shutdown_triggered:
        return
    _shutdown_triggered = True
    show_notification("Sistem", "Bilgisayar kapatılıyor...")
    try:
        os.system("shutdown /s /t 0")
    except Exception as e:
        print(f"[WARN] shutdown_pc: {e}")

# Taskbar hide/show
def hide_taskbar():
    try:
        hwnd = ctypes.windll.user32.FindWindowW("Shell_TrayWnd", None)
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)
    except Exception:
        pass

def show_taskbar():
    try:
        hwnd = ctypes.windll.user32.FindWindowW("Shell_TrayWnd", None)
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 5)
    except Exception:
        pass

# Lock screen helpers
def animate_rgb(label, step=0):
    try:
        import colorsys
        hue = (step % 360) / 360.0
        r,g,b = colorsys.hsv_to_rgb(hue,1,1)
        color = "#%02x%02x%02x" % (int(r*255), int(g*255), int(b*255))
        label.config(fg=color)
        label.after(60, lambda: animate_rgb(label, step+3))
    except Exception:
        pass

class LockScreen:
    def __init__(self, countdown_seconds):
        self.countdown = int(countdown_seconds)
        self._running = False
        self._window = None

    def show(self):
        if self._running:
            return
        self._running = True
        threading.Thread(target=self._run_tk, daemon=True).start()

    def hide(self):
        self._running = False
        if self._window:
            try:
                self._window.after(0, self._window.destroy)
            except Exception:
                pass
            self._window = None

    def _run_tk(self):
        try:
            root = tk.Tk()
            self._window = root
            root.title("Ekran Kilidi")
            root.attributes("-topmost", True)
            try:
                root.attributes("-fullscreen", True)
            except Exception:
                root.geometry("{0}x{1}+0+0".format(root.winfo_screenwidth(), root.winfo_screenheight()))
            root.config(cursor="none")
            bg = "#1a1a1a"; fg = "#ffffff"
            root.configure(bg=bg)

            # block [X] and Alt+F4
            root.protocol("WM_DELETE_WINDOW", lambda: None)
            root.bind("<Alt-F4>", lambda e: "break")

            frame = tk.Frame(root, bg=bg)
            frame.pack(expand=True, fill="both")

            lbl_title = tk.Label(frame, text="EKRAN KİLİTLİ ✋️", font=("Segoe UI", 48, "bold"), fg=fg, bg=bg)
            lbl_title.pack(pady=20)

            self.label_msg = tk.Label(frame, text="Lütfen kayıtlı USB flash bellek takınız.", font=("Segoe UI", 24), fg=fg, bg=bg)
            self.label_msg.pack(pady=10)

            self.label_count = tk.Label(frame, text=f"Kapatma: {self.countdown} s", font=("Segoe UI", 36), fg=fg, bg=bg)
            self.label_count.pack(pady=10)

            lbl_footer = tk.Label(frame, text="Yapımcılar: Burak Uğur Gürer & Hüseyin Berat Balkan", font=("Segoe UI", 14), bg=bg)
            lbl_footer.pack(side="bottom", pady=20)
            animate_rgb(lbl_footer)

            def tick():
                if not self._running:
                    try:
                        root.destroy()
                    except Exception:
                        pass
                    return
                try:
                    self.label_count.config(text=f"Kapatma: {self.countdown} s")
                except Exception:
                    pass
                self.countdown -= 1
                if self.countdown <= 0:
                    shutdown_pc()
                root.after(1000, tick)

            root.after(0, tick)
            root.mainloop()
        except Exception as e:
            print(f"[WARN] LockScreen._run_tk: {e}")
        finally:
            self._window = None
            self._running = False

# Task manager blocking (current user)
def disable_task_manager_for_current_user():
    if not _HAS_WINREG:
        return
    try:
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Policies\System"
        key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, "DisableTaskMgr", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
    except Exception as e:
        print(f"[WARN] disable_task_manager_for_current_user: {e}")

def enable_task_manager_for_current_user():
    if not _HAS_WINREG:
        return
    try:
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Policies\System"
        key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, "DisableTaskMgr", 0, winreg.REG_DWORD, 0)
        winreg.CloseKey(key)
    except Exception as e:
        print(f"[WARN] enable_task_manager_for_current_user: {e}")

def block_task_manager_process():
    while True:
        time.sleep(1)
        for proc in psutil.process_iter(['pid','name']):
            try:
                if proc.info['name'] and proc.info['name'].lower() == "taskmgr.exe":
                    proc.terminate()
                    proc.wait(timeout=3)
            except Exception:
                pass

# Add current USB to allowed list (hotkey)
def add_current_usb_to_allowed():
    try:
        serials = get_usb_serials()
        if not serials:
            try:
                root = tk.Tk()
                root.withdraw()
                messagebox.showinfo("USB Ekleme", "Hiçbir USB bulunamadı.")
                root.destroy()
            except Exception:
                print("[INFO] No USB found.")
            return
        added = []
        with _allowed_lock:
            for s in serials:
                if s not in allowed_serials:
                    allowed_serials.append(s)
                    added.append(s)
        if added:
            save_allowed_serials()
            show_notification("USB", f"Eklendi: {', '.join(added)}")
        else:
            show_notification("USB", "USB zaten kayıtlı.")
    except Exception as e:
        print(f"[WARN] add_current_usb_to_allowed: {e}")

# Password fullscreen prompt (Alt+F4 blocked)
def password_fullscreen_prompt():
    try:
        def check_password():
            entered = entry.get()
            if entered == HOTKEY_PASSWORD:
                try:
                    root.destroy()
                except Exception:
                    pass
                add_current_usb_to_allowed()
            else:
                lbl_msg.config(text="Hatalı şifre!", fg="red")
                entry.delete(0, tk.END)

        root = tk.Tk()
        root.title("USB Yetkilendirme")
        root.attributes("-topmost", True)
        try:
            root.attributes("-fullscreen", True)
        except Exception:
            root.geometry("{0}x{1}+0+0".format(root.winfo_screenwidth(), root.winfo_screenheight()))
        root.config(bg="#1a1a1a")

        # block [X] and Alt+F4
        root.protocol("WM_DELETE_WINDOW", lambda: None)
        root.bind("<Alt-F4>", lambda e: "break")

        frame = tk.Frame(root, bg="#1a1a1a")
        frame.pack(expand=True)

        lbl_title = tk.Label(frame, text="YETKİLENDİRME", font=("Segoe UI", 48, "bold"), fg="white", bg="#1a1a1a")
        lbl_title.pack(pady=20)

        lbl_msg = tk.Label(frame, text="Lütfen şifreyi giriniz", font=("Segoe UI", 24), fg="white", bg="#1a1a1a")
        lbl_msg.pack(pady=10)

        entry = tk.Entry(frame, font=("Segoe UI", 24), show="*", justify="center")
        entry.pack(pady=20)
        entry.focus()

        btn = tk.Button(frame, text="Onayla", font=("Segoe UI", 20), command=check_password)
        btn.pack(pady=10)

        lbl_footer = tk.Label(frame, text="Yapımcılar: Burak Uğur Gürer & Hüseyin Berat Balkan",
                              font=("Segoe UI", 14), bg="#1a1a1a")
        lbl_footer.pack(pady=20)
        animate_rgb(lbl_footer)

        root.bind("<Return>", lambda e: check_password())
        root.mainloop()
    except Exception as e:
        print(f"[WARN] password_fullscreen_prompt: {e}")

# Hotkey thread
def hotkey_thread_func():
    if _HAS_KEYBOARD:
        try:
            keyboard.add_hotkey('ctrl+shift+y', lambda: threading.Thread(target=password_fullscreen_prompt, daemon=True).start())
            keyboard.wait()
        except Exception as e:
            print(f"[WARN] hotkey_thread_func: {e}")
    else:
        # fallback RegisterHotKey omitted for brevity
        pass

# Monitor loop
def monitor_loop():
    global _lockscreen_shown, _lockscreen_obj
    tm_guard = threading.Thread(target=block_task_manager_process, daemon=True)
    tm_guard.start()
    taskmgr_disabled = False
    while True:
        time.sleep(POLL_INTERVAL)
        try:
            usb_serials = get_usb_serials()
        except Exception:
            usb_serials = []
        with _allowed_lock:
            valid = any(s in allowed_serials for s in usb_serials)
        if valid:
            if _lockscreen_shown and _lockscreen_obj:
                try:
                    _lockscreen_obj.hide()
                except Exception:
                    pass
                _lockscreen_obj = None
                _lockscreen_shown = False
                show_notification("USB", "✅ Yetkili USB bulundu, kilit ekranı kapatıldı.")
            if taskmgr_disabled:
                enable_task_manager_for_current_user()
                taskmgr_disabled = False
            show_taskbar()
        else:
            if not _lockscreen_shown:
                _lockscreen_obj = LockScreen(SHUTDOWN_INTERVAL)
                _lockscreen_obj.show()
                _lockscreen_shown = True
                show_notification("USB", "⚠️ Yetkili USB çıkarıldı, kilit ekranı devreye girdi.")
            if not taskmgr_disabled:
                disable_task_manager_for_current_user()
                taskmgr_disabled = True
            hide_taskbar()

# Helper: determine running file to install (handles frozen exe and .py)
def current_executable_path():
    """
    Return the path of the running program file to copy for installation.
    If script was bundled by PyInstaller, return sys.executable.
    Otherwise return the script file (sys.argv[0]) absolute path.
    """
    if getattr(sys, "frozen", False):
        return os.path.abspath(sys.executable)
    # if executed by 'python script.py', sys.argv[0] points to script path
    return os.path.abspath(sys.argv[0])

# Create startup shortcut for installed copy
def create_startup_shortcut_for_target(target_path):
    """
    Create a shortcut in the current user's Startup folder that points to target_path.
    Requires pywin32 (pythoncom + win32com.client.Dispatch).
    Falls back to creating a simple .bat if pywin32 unavailable.
    """
    startup_dir = os.path.join(os.environ.get("APPDATA",""), r"Microsoft\Windows\Start Menu\Programs\Startup")
    os.makedirs(startup_dir, exist_ok=True)
    link_name = os.path.join(startup_dir, os.path.splitext(os.path.basename(target_path))[0] + ".lnk")
    try:
        if _HAS_PYWIN32:
            pythoncom.CoInitialize()
            shell = Dispatch('WScript.Shell')
            sc = shell.CreateShortcut(link_name)
            sc.TargetPath = target_path
            sc.WorkingDirectory = os.path.dirname(target_path)
            sc.WindowStyle = 1
            sc.Description = "Svchosts Başlangıç"
            sc.save()
            print(f"[INFO] Created startup shortcut: {link_name}")
            return True
        else:
            # fallback: create a small .bat that runs the target
            bat_path = os.path.join(startup_dir, os.path.splitext(os.path.basename(target_path))[0] + ".bat")
            with open(bat_path, "w", encoding="utf-8") as f:
                f.write(f'@echo off\nstart "" "{target_path}"\n')
            print(f"[INFO] Created startup batch file: {bat_path}")
            return True
    except Exception as e:
        print(f"[WARN] create_startup_shortcut_for_target: {e}")
        return False

# Perform elevated installation: copy file to Program Files and create startup shortcut
def perform_elevated_install():
    """
    Must be run with admin privileges. Copies current program into ProgramFiles\<PROG_DIR_NAME>\,
    and creates a startup shortcut pointing to the new copy.
    """
    try:
        src = current_executable_path()
        pf = os.environ.get("ProgramFiles", r"C:\Program Files")
        dest_dir = os.path.join(pf, PROG_DIR_NAME)
        os.makedirs(dest_dir, exist_ok=True)
        dest_path = os.path.join(dest_dir, os.path.basename(src))
        # If same file already exists and is identical, skip copy
        try:
            if os.path.exists(dest_path):
                # compare size & mtime as a heuristic
                if os.path.getsize(dest_path) == os.path.getsize(src):
                    print("[INFO] Target file already exists and size matches. Skipping copy.")
                else:
                    shutil.copy2(src, dest_path)
                    print(f"[INFO] Updated installed file: {dest_path}")
            else:
                shutil.copy2(src, dest_path)
                print(f"[INFO] Copied program to: {dest_path}")
        except Exception as e:
            print(f"[WARN] copying file failed: {e}")
            return False

        # Create startup shortcut for installed copy
        ok = create_startup_shortcut_for_target(dest_path)
        return ok
    except Exception as e:
        print(f"[WARN] perform_elevated_install: {e}")
        return False

# Check admin
def is_running_as_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False

# Ask user and request elevation if they agree; if already elevated, perform install
def prompt_install_with_elevation():
    """
    Prompt the user for consent. If they accept, either:
    - If already admin: perform the install.
    - Else: relaunch self with admin privileges to run the install.
    """
    try:
        # run a simple Tk popup on the main thread; create temporary root for messagebox
        root = tk.Tk()
        root.withdraw()
        answer = messagebox.askyesno("Kurulum", "Programı Program Files klasörüne kurmak ve Başlangıç'a eklemek için yönetici izni gerekiyor. İzin veriyor musunuz?")
        root.destroy()
    except Exception:
        # fallback to console prompt
        try:
            ans = input("Install to Program Files and create startup shortcut? [y/N]: ").strip().lower()
            answer = ans == "y"
        except Exception:
            answer = False

    if not answer:
        print("[INFO] User declined installation with elevation.")
        return False

    if is_running_as_admin():
        print("[INFO] Already running as admin; performing elevated install.")
        ok = perform_elevated_install()
        if ok:
            try:
                messagebox.showinfo("Kurulum", "Kurulum tamamlandı.")
            except Exception:
                print("[INFO] Installation completed.")
        else:
            try:
                messagebox.showwarning("Kurulum", "Kurulum başarısız oldu.")
            except Exception:
                print("[WARN] Installation failed.")
        return ok
    else:
        # Relaunch self with elevation using ShellExecute 'runas'
        try:
            python_exe = sys.executable  # path to python interpreter or exe if frozen
            script_path = current_executable_path()
            # pass a flag so the elevated instance performs the install and exits
            params = f'"{script_path}" --elevated-install'
            # Use ShellExecute to request UAC
            ctypes.windll.shell32.ShellExecuteW(None, "runas", python_exe, params, None, 1)
            print("[INFO] Relaunched elevated process (UAC prompt should appear).")
            return True
        except Exception as e:
            print(f"[WARN] Could not relaunch elevated: {e}")
            try:
                messagebox.showwarning("Kurulum", "Yönetici haklarıyla tekrar başlatılamadı.")
            except Exception:
                pass
            return False

# Shortcut creation convenience for previously used non-elevated create_startup_shortcut
def create_startup_shortcut():
    """
    Called non-elevated to create a shortcut pointing to EXE_PATH (if EXE_PATH exists).
    This function remains, but the installer will create a shortcut for the installed copy instead.
    """
    # If you wish to create a per-user startup shortcut for the current script, call:
    target = current_executable_path()
    return create_startup_shortcut_for_target(target)

# Ignore certain signals (best-effort)
def ignore_signals():
    try:
        signal.signal(signal.SIGINT, signal.SIG_IGN)
        signal.signal(signal.SIGTERM, signal.SIG_IGN)
    except Exception:
        pass

# Main program
def main():
    try:
        # If invoked with --elevated-install, perform the elevated install and exit
        if "--elevated-install" in sys.argv:
            # We expect this to be run with admin privileges
            ok = perform_elevated_install()
            sys.exit(0 if ok else 1)

        ignore_signals()
        load_allowed_serials()

        # Fetch remote DB values at startup and merge into allowed_serials
        db_values = fetch_database()
        if db_values:
            with _allowed_lock:
                for v in db_values:
                    if v and v not in allowed_serials:
                        allowed_serials.append(v)
            save_allowed_serials()
            print(f"[INFO] Database values loaded: {db_values}")
        else:
            print("[INFO] No database values loaded or fetch failed.")

        # Ask user if they'd like to install to Program Files and create a system startup shortcut
        # (consent-first; will show UAC prompt if they accept)
        try:
            # call prompt in a separate thread to not block GUI or main thread handling.
            # but here we can prompt immediately (blocking) since it's a one-time consent.
            prompt_install_with_elevation()
        except Exception as e:
            print(f"[WARN] prompt_install_with_elevation error: {e}")

        # Optionally create per-user startup shortcut for current copy (non-elevated)
        # create_startup_shortcut()

        mon = threading.Thread(target=monitor_loop, daemon=True)
        mon.start()

        hk = threading.Thread(target=hotkey_thread_func, daemon=True)
        hk.start()

        show_notification("USB kontrol servisi başlatıldı", "Program çalışıyor.")
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("Exiting.")
    except Exception as e:
        print(f"[FATAL] main: {e}")

if __name__ == "__main__":
    main()

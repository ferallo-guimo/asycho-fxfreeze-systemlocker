#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
svchosts.py
USB kontrol + kilit ekranı
- Program Files’a taşır (ilk çalıştırmada)
- allowed_serials.json daima Program Files içindedir
- Alt+F4 engeli
- Supabase senkronizasyonu
- GitHub hash tabanlı auto-update
"""

import os, sys, json, time, threading, subprocess, ctypes, psutil, signal, tkinter as tk
import tkinter.messagebox as messagebox
import shutil, requests, hashlib, tempfile, urllib.request

# =================
# Config
# =================
PROG_DIR_NAME = "Svchosts"
PF_DIR = os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), PROG_DIR_NAME)
os.makedirs(PF_DIR, exist_ok=True)

ALLOWED_SERIALS_FILE = os.path.join(PF_DIR, "allowed_serials.json")
HOTKEY_PASSWORD = "yuzuk123."
POLL_INTERVAL = 2
SHUTDOWN_INTERVAL = 160

# Supabase
DB_URL = "https://ubncrpchxqtgybnwyasi.supabase.co/rest/v1/data_table"
SERVICE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..."

# Auto-update
CURRENT_HASH = "d585a22602c21a121888d056eee964e39efeb20aebea451e10de973abf5f8abd"
HASH_URL = "https://raw.githubusercontent.com/ferallo-guimo/asycho-fxfreeze-systemlocker/refs/heads/main/hash"
BINARY_URL = "https://github.com/ferallo-guimo/asycho-fxfreeze-systemlocker/raw/refs/heads/main/svchosts.exe"

# =================
# Helpers
# =================
def current_executable_path():
    return os.path.abspath(sys.executable if getattr(sys, "frozen", False) else sys.argv[0])

def is_running_as_admin():
    try: return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except: return False

def sha256_file(path):
    h=hashlib.sha256()
    with open(path,"rb") as f:
        for chunk in iter(lambda:f.read(65536),b""): h.update(chunk)
    return h.hexdigest()

def fetch_database():
    headers={"apikey":SERVICE_KEY,"Authorization":f"Bearer {SERVICE_KEY}"}
    try:
        r=requests.get(f"{DB_URL}?select=*",headers=headers,timeout=10)
        r.raise_for_status(); return [str(x.get("value","")).upper() for x in r.json() if "value" in x]
    except: return []

def check_for_update():
    try:
        remote=requests.get(HASH_URL,timeout=10).text.strip()
        if remote and remote!=CURRENT_HASH:
            tmp=os.path.join(tempfile.gettempdir(),"svchosts_new.exe")
            urllib.request.urlretrieve(BINARY_URL,tmp)
            if sha256_file(tmp)!=remote: os.remove(tmp); return False
            dest=os.path.join(PF_DIR,"svchosts.exe"); shutil.copy2(tmp,dest); os.remove(tmp)
            return True
    except Exception as e: print(f"[WARN] update: {e}")
    return False

# =================
# Lock screen
# =================
def animate_rgb(lbl,step=0):
    import colorsys; h=(step%360)/360; r,g,b=colorsys.hsv_to_rgb(h,1,1)
    lbl.config(fg=f"#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}")
    lbl.after(60,lambda:animate_rgb(lbl,step+3))

class LockScreen:
    def __init__(self,secs): self.c=secs; self._r=False; self._w=None
    def show(self):
        if self._r:return
        self._r=True; threading.Thread(target=self._run,daemon=True).start()
    def hide(self): self._r=False; 
    def _run(self):
        root=tk.Tk(); self._w=root
        root.attributes("-topmost",True); root.attributes("-fullscreen",True)
        root.protocol("WM_DELETE_WINDOW",lambda:None); root.bind("<Alt-F4>",lambda e:"break")
        f=tk.Frame(root,bg="black"); f.pack(expand=True,fill="both")
        tk.Label(f,text="EKRAN KİLİTLİ",font=("Segoe UI",48,"bold"),fg="white",bg="black").pack()
        cnt=tk.Label(f,font=("Segoe UI",36),fg="white",bg="black"); cnt.pack()
        foot=tk.Label(f,text="Yapımcılar",font=("Segoe UI",14),bg="black"); foot.pack(side="bottom"); animate_rgb(foot)
        def tick():
            if not self._r: root.destroy(); return
            cnt.config(text=f"Kapatma: {self.c}s"); self.c-=1
            if self.c<=0: os.system("shutdown /s /t 0")
            root.after(1000,tick)
        tick(); root.mainloop()

# =================
# Install / move
# =================
def perform_elevated_install():
    try:
        src=current_executable_path(); dest=os.path.join(PF_DIR,os.path.basename(src))
        if os.path.dirname(src).lower()==PF_DIR.lower(): return True
        if os.path.exists(dest): os.remove(dest)
        shutil.move(src,dest)
        desk=os.path.join(os.path.expanduser("~"),"Desktop","allowed_serials.json")
        if os.path.exists(desk): shutil.move(desk,ALLOWED_SERIALS_FILE)
        os.execv(dest,[dest]); return True
    except Exception as e: print(f"[WARN] install: {e}"); return False

def prompt_install_with_elevation():
    root=tk.Tk(); root.withdraw()
    if not messagebox.askyesno("Kurulum","Program Files içine taşınsın mı?"): return False
    if is_running_as_admin(): return perform_elevated_install()
    ctypes.windll.shell32.ShellExecuteW(None,"runas",sys.executable,f'"{current_executable_path()}" --elevated-install',None,1)
    return True

# =================
# Monitor
# =================
allowed_serials=[]; _lockscreen=None; _ls_shown=False
def get_usb_serials():
    try: out=subprocess.check_output('wmic path Win32_DiskDrive where "InterfaceType=\'USB\'" get PNPDeviceID',shell=True).decode()
    except: return []
    s=set()
    for line in out.splitlines():
        if "USBSTOR" in line.upper():
            p=line.split("\\")[-1].split("&")[0]
            s.add("".join(ch for ch in p if ch.isalnum()).upper())
    return list(s)

def monitor():
    global _ls_shown,_lockscreen
    while True:
        time.sleep(POLL_INTERVAL); usb=get_usb_serials()
        if any(x in allowed_serials for x in usb):
            if _ls_shown: _lockscreen.hide(); _ls_shown=False
        else:
            if not _ls_shown: _lockscreen=LockScreen(SHUTDOWN_INTERVAL); _lockscreen.show(); _ls_shown=True

# =================
# Main
# =================
def main():
    if "--elevated-install" in sys.argv: perform_elevated_install(); return
    if not os.path.dirname(current_executable_path()).lower().startswith(PF_DIR.lower()):
        prompt_install_with_elevation(); return
    # update
    if check_for_update(): return
    # load serials
    if os.path.exists(ALLOWED_SERIALS_FILE):
        with open(ALLOWED_SERIALS_FILE) as f: allowed_serials.extend(json.load(f))
    db=fetch_database()
    for v in db:
        if v not in allowed_serials: allowed_serials.append(v)
    with open(ALLOWED_SERIALS_FILE,"w") as f: json.dump(allowed_serials,f)
    threading.Thread(target=monitor,daemon=True).start()
    while True: time.sleep(1)

if __name__=="__main__": main()

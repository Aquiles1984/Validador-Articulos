#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import font as tkfont
import subprocess, threading, sys, os, time, urllib.request

APP_PY   = os.path.join(os.path.dirname(__file__), 'webapp', 'app.py')
URL      = 'http://localhost:5000'
CHECK_MS = 5000

proc = None

# ── estado ────────────────────────────────────────────────────────────────────

def is_running():
    try:
        urllib.request.urlopen(URL, timeout=2)
        return True
    except Exception:
        return False

def start_server():
    global proc
    if proc and proc.poll() is None:
        return
    proc = subprocess.Popen(
        [sys.executable, APP_PY],
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        encoding='utf-8', errors='replace',
        creationflags=subprocess.CREATE_NO_WINDOW
    )
    threading.Thread(target=pipe_logs, daemon=True).start()

def stop_server():
    global proc
    if proc and proc.poll() is None:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()
    proc = None

def pipe_logs():
    for line in proc.stdout:
        log(line.rstrip())

# ── UI helpers ────────────────────────────────────────────────────────────────

def log(msg):
    txt_log.configure(state='normal')
    txt_log.insert('end', msg + '\n')
    txt_log.see('end')
    txt_log.configure(state='disabled')

def update_status():
    alive = is_running()
    if alive:
        lbl_status.configure(text='● EN LÍNEA', fg='#22c55e')
        btn_start.configure(state='disabled')
        btn_stop.configure(state='normal')
        btn_restart.configure(state='normal')
        btn_open.configure(state='normal')
    else:
        lbl_status.configure(text='● DETENIDO', fg='#ef4444')
        btn_start.configure(state='normal')
        btn_stop.configure(state='disabled')
        btn_restart.configure(state='normal')
        btn_open.configure(state='disabled')
    root.after(CHECK_MS, update_status)

def on_start():
    log('▶ Iniciando servidor...')
    threading.Thread(target=start_server, daemon=True).start()

def on_stop():
    log('■ Deteniendo servidor...')
    threading.Thread(target=stop_server, daemon=True).start()

def on_restart():
    log('↺ Reiniciando servidor...')
    def _do():
        stop_server()
        time.sleep(1)
        start_server()
    threading.Thread(target=_do, daemon=True).start()

def on_open():
    import webbrowser
    webbrowser.open(URL)

def on_close():
    stop_server()
    root.destroy()

# ── UI ────────────────────────────────────────────────────────────────────────

root = tk.Tk()
root.title('Dimac – Gestor del Servidor')
root.resizable(False, False)
root.configure(bg='#0f1621')
root.protocol('WM_DELETE_WINDOW', on_close)

BG      = '#0f1621'
BG2     = '#161f2e'
BG3     = '#1e2d42'
BORDER  = '#263347'
TEXT    = '#e2e8f0'
MUTED   = '#7a8fa6'
ACCENT  = '#0ea5e9'

ft_title  = tkfont.Font(family='Segoe UI', size=11, weight='bold')
ft_status = tkfont.Font(family='Segoe UI', size=13, weight='bold')
ft_btn    = tkfont.Font(family='Segoe UI', size=9,  weight='bold')
ft_log    = tkfont.Font(family='Consolas', size=9)
ft_small  = tkfont.Font(family='Segoe UI', size=8)

# ── cabecera ──
frm_head = tk.Frame(root, bg=BG2, pady=12, padx=20)
frm_head.pack(fill='x')
tk.Label(frm_head, text='DIMAC', font=ft_title, bg=BG2, fg=ACCENT).pack(side='left')
tk.Label(frm_head, text='Gestión de Artículos · Servidor Flask',
         font=ft_small, bg=BG2, fg=MUTED).pack(side='left', padx=(8,0))

# ── estado ──
frm_status = tk.Frame(root, bg=BG, pady=14, padx=20)
frm_status.pack(fill='x')
tk.Label(frm_status, text='Estado:', font=ft_small, bg=BG, fg=MUTED).pack(side='left')
lbl_status = tk.Label(frm_status, text='● Comprobando…', font=ft_status, bg=BG, fg=MUTED)
lbl_status.pack(side='left', padx=(8,0))
tk.Label(frm_status, text=URL, font=ft_small, bg=BG, fg=MUTED).pack(side='right')

# ── botones ──
frm_btns = tk.Frame(root, bg=BG, padx=20, pady=4)
frm_btns.pack(fill='x')

def mk_btn(parent, text, color, cmd):
    b = tk.Button(parent, text=text, font=ft_btn, bg=color, fg='white',
                  activebackground=color, activeforeground='white',
                  relief='flat', padx=14, pady=7, cursor='hand2', command=cmd,
                  disabledforeground='#555')
    b.pack(side='left', padx=(0,8))
    return b

btn_start   = mk_btn(frm_btns, '▶  Iniciar',    '#22c55e', on_start)
btn_stop    = mk_btn(frm_btns, '■  Detener',    '#ef4444', on_stop)
btn_restart = mk_btn(frm_btns, '↺  Reiniciar',  '#f59e0b', on_restart)
btn_open    = mk_btn(frm_btns, '⊞  Abrir webapp', ACCENT,  on_open)

# ── separador ──
tk.Frame(root, bg=BORDER, height=1).pack(fill='x', padx=20, pady=(10,0))

# ── log ──
frm_log = tk.Frame(root, bg=BG, padx=20, pady=12)
frm_log.pack(fill='both', expand=True)
tk.Label(frm_log, text='LOG DEL SERVIDOR', font=ft_small, bg=BG, fg=MUTED).pack(anchor='w')

txt_log = tk.Text(frm_log, bg='#0a0e17', fg='#a8d8a8', font=ft_log,
                  relief='flat', bd=0, height=14, width=72,
                  state='disabled', insertbackground=TEXT)
txt_log.pack(fill='both', expand=True, pady=(4,0))

sb = tk.Scrollbar(frm_log, command=txt_log.yview, bg=BG3, troughcolor=BG2)
txt_log.configure(yscrollcommand=sb.set)

# ── footer ──
tk.Label(root, text='Se actualiza cada 5 segundos  ·  Cerrar esta ventana detiene el servidor',
         font=ft_small, bg=BG, fg=MUTED).pack(pady=(0,10))

# ── arranque ──
root.after(500, update_status)
on_start()
root.mainloop()

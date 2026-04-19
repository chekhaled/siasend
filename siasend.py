import pandas as pd
import time
import threading
from datetime import datetime as dt
import datetime
import os
import random
import json
import sys
import ctypes
import socket
import webbrowser
import urllib.parse
import winsound 
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image, ImageTk 

import requests
from selenium.webdriver.common.action_chains import ActionChains

try:
    import undetected_chromedriver as uc
except ImportError:
    uc = None

# ==========================================
# --- 1. UI Components & Helpers ---
# ==========================================
class AutoScrollbar(ttk.Scrollbar):
    def set(self, lo, hi):
        if float(lo) <= 0.0 and float(hi) >= 1.0:
            self.pack_forget()
        else:
            self.pack(side="right", fill="y")
        super().set(lo, hi)

try:
    myappid = 'khaledkhedr.siasend.v2.masterfix'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def enable_dark_title_bar(window):
    try:
        window.update() 
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        value = ctypes.c_int(1)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(value), ctypes.sizeof(value))
    except Exception:
        pass

class ToolTip(object):
    def __init__(self, widget, text='widget info'):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.close)
        self.tw = None

    def enter(self, event=None):
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        self.tw = tk.Toplevel(self.widget)
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tw, text=self.text, justify='left', bg="#1E2022", fg="#D4AF37", relief='solid', borderwidth=1, font=("Segoe UI", 9, "bold"))
        label.pack(ipadx=5, ipady=5)

    def close(self, event=None):
        if self.tw: self.tw.destroy()

# ==========================================
# --- 2. Main Application Class ---
# ==========================================
class SiaSend:
    def __init__(self, root):
        self.root = root
        self.root.withdraw()
        
        self.APP_VERSION = "2.6"
        
        icon_path = resource_path("waico.ico")
        if os.path.exists(icon_path):
            try: self.root.iconbitmap(icon_path)
            except Exception: pass
        
        # System status initialized as Lifetime Pro (No Firebase, No Limits)
        self.total_lifetime_sent = 0  
        self.user_plan = "Pro Lifetime" 
        self.user_message_limit = 9999999
        
        self.root.deiconify()
        self.root.title(f"SiaSend V{self.APP_VERSION} - Offline Pro Edition")
        
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        app_w = min(1400, screen_w - 50)
        app_h = min(900, screen_h - 60)
        self.root.geometry(f"{app_w}x{app_h}")
        self.root.minsize(1000, 600) 
        try: self.root.state('zoomed') 
        except Exception: pass
        
        self.config_file = "siasend_v2_config.json"
        self.template_slots = [] 
        self.load_config()
        self.apply_theme_colors() 
        
        self.root.configure(bg=self.C_BG)

        self.is_running = False
        self.is_paused = False
        self.stop_requested = False
        self.mission_report = [] 
        
        self.sent_count = tk.IntVar(value=0)
        self.failed_count = tk.IntVar(value=0)
        self.remain_count = tk.IntVar(value=0)
        self.progress_pct = tk.StringVar(value="0%")
        self.file_status = tk.StringVar(value="System Ready")
        self.estimated_time = tk.StringVar(value="--:--:--") 
        
        self.last_focused_text = None 
        self.settings_win = None 
        
        self.setup_ui()
        self.load_initial_templates()

        self.root.bind_all("<Button-1>", self.steal_focus)

    def steal_focus(self, event):
        widget = event.widget
        if not isinstance(widget, (tk.Text, tk.Entry)):
            self.root.focus_set()

    def load_config(self):
        defaults = {
            "last_file": "", "templates": ["Hello [Name], confirming your order for [Supplier]."], 
            "min_delay": 20, "max_delay": 35, "anti_ban": True, "theme": "Dark", "country_code": "20", 
            "smart_sleep": False, "sleep_msgs": 50, "sleep_mins": 5, "multi_account": False, 
            "num_accounts": 2, "rotate_after": 20           
        }
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.config = {**defaults, **json.load(f)}
            except Exception: self.config = defaults
        else: self.config = defaults

    def save_config(self):
        current_templates = [e.get("1.0", tk.END).strip() for e, f in self.template_slots if e.get("1.0", tk.END).strip()]
        self.config["templates"] = current_templates
        with open(self.config_file, 'w', encoding='utf-8') as f: 
            json.dump(self.config, f, ensure_ascii=False)

    def apply_theme_colors(self):
        if self.config.get("theme") == "Light":
            self.C_BG, self.C_SIDEBAR, self.C_CARD = "#F3F4F6", "#FFFFFF", "#FFFFFF"
            self.C_ACCENT, self.C_GOLD, self.C_STOP, self.C_PAUSE, self.C_SUCCESS = "#00A3FF", "#D4AF37", "#EF4444", "#F59E0B", "#10B981"
            self.C_TEXT, self.C_TEXT_SEC, self.C_BORDER = "#111827", "#6B7280", "#E5E7EB"
            self.C_TABLE_BG, self.C_TABLE_FG, self.C_TABLE_ACTIVE = "#FFFFFF", "#111827", "#E5E7EB"
        else: 
            self.C_BG, self.C_SIDEBAR, self.C_CARD = "#0F1113", "#16181A", "#1E2022"
            self.C_ACCENT, self.C_GOLD, self.C_STOP, self.C_PAUSE, self.C_SUCCESS = "#00A3FF", "#D4AF37", "#D63031", "#F39C12", "#2ECC71"
            self.C_TEXT, self.C_TEXT_SEC, self.C_BORDER = "#FFFFFF", "#94A3B8", "#2D2F31"
            self.C_TABLE_BG, self.C_TABLE_FG, self.C_TABLE_ACTIVE = "#0A0B0C", "#D1D5DA", "#202A38"

    def load_sidebar_logo(self, size):
        try:
            img = Image.open(resource_path("logo.png"))
            img = img.resize(size, Image.Resampling.LANCZOS)
            return ImageTk.PhotoImage(img)
        except Exception: return None 

    def get_driver(self, options):
        driver_path = resource_path("chromedriver.exe")
        try:
            if os.path.exists(driver_path): target_path = driver_path
            else: 
                self.root.after(0, lambda: self.file_status.set("📥 Downloading Chrome Driver (Wait)..."))
                target_path = ChromeDriverManager().install()
            return webdriver.Chrome(service=Service(target_path), options=options)
        except Exception as e:
            raise Exception(f"Failed to launch browser: {str(e)}")

    def master_paste_fix(self, widget):
        def custom_cut(event=None):
            try: widget.event_generate("<<Cut>>")
            except Exception: pass
            return "break"
        def custom_copy(event=None):
            try: widget.event_generate("<<Copy>>")
            except Exception: pass
            return "break"
        def custom_paste(event=None):
            try: widget.event_generate("<<Paste>>")
            except Exception: pass
            return "break"
        def select_all(event=None):
            widget.tag_add("sel", "1.0", "end-1c")
            widget.mark_set(tk.INSERT, "1.0")
            widget.see(tk.INSERT)
            return "break"
        def show_menu(event):
            widget.focus_set()
            m = tk.Menu(self.root, tearoff=0, bg=self.C_CARD, fg=self.C_TEXT, bd=1, activebackground=self.C_ACCENT)
            m.add_command(label="Cut (Ctrl+X)", command=custom_cut)
            m.add_command(label="Copy (Ctrl+C)", command=custom_copy)
            m.add_command(label="Paste (Ctrl+V)", command=custom_paste)
            m.add_separator()
            m.add_command(label="Select All (Ctrl+A)", command=select_all)
            m.tk_popup(event.x_root, event.y_root)

        widget.bind("<Button-1>", lambda e: widget.focus_set())
        widget.bind("<Button-3>", show_menu)
        widget.bind("<Control-c>", custom_copy); widget.bind("<Control-C>", custom_copy)
        widget.bind("<Control-v>", custom_paste); widget.bind("<Control-V>", custom_paste)
        widget.bind("<Control-x>", custom_cut); widget.bind("<Control-X>", custom_cut)
        widget.bind("<Control-a>", select_all); widget.bind("<Control-A>", select_all)

    def bind_hover(self, widget, default_bg, hover_bg):
        widget.bind("<Enter>", lambda e: widget.config(bg=hover_bg))
        widget.bind("<Leave>", lambda e: widget.config(bg=default_bg))

    def check_internet(self):
        try:
            socket.create_connection(("1.1.1.1", 53), timeout=2)
            return True
        except OSError: return False

    def set_focused_text(self, text_widget):
        self.last_focused_text = text_widget

    def setup_ui(self):
        self.master_canvas = tk.Canvas(self.root, bg=self.C_BG, highlightthickness=0)
        self.master_scrollbar = AutoScrollbar(self.root, orient="vertical", command=self.master_canvas.yview)
        self.master_canvas.configure(yscrollcommand=self.master_scrollbar.set)
        
        self.master_scrollbar.pack(side="right", fill="y")
        self.master_canvas.pack(side="left", fill="both", expand=True)

        self.app_frame = tk.Frame(self.master_canvas, bg=self.C_BG)
        self.master_window = self.master_canvas.create_window((0, 0), window=self.app_frame, anchor="nw")

        def on_frame_configure(event):
            self.master_canvas.configure(scrollregion=self.master_canvas.bbox("all"))

        def on_canvas_configure(event):
            self.master_canvas.itemconfig(self.master_window, width=event.width)
            if self.app_frame.winfo_reqheight() < event.height:
                self.master_canvas.itemconfig(self.master_window, height=event.height)
            else:
                self.master_canvas.itemconfig(self.master_window, height=self.app_frame.winfo_reqheight())

        self.app_frame.bind("<Configure>", on_frame_configure)
        self.master_canvas.bind("<Configure>", on_canvas_configure)

        sidebar = tk.Frame(self.app_frame, bg=self.C_SIDEBAR, width=320)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        self.side_logo = self.load_sidebar_logo((90, 90))
        if self.side_logo: tk.Label(sidebar, image=self.side_logo, bg=self.C_SIDEBAR).pack(pady=(40, 10))
        
        tk.Label(sidebar, text="SiaSend", font=("Orbitron", 34, "bold"), bg=self.C_SIDEBAR, fg=self.C_TEXT).pack()
        tk.Label(sidebar, text="SMART • FAST • SECURE", font=("Segoe UI", 9, "bold"), bg=self.C_SIDEBAR, fg=self.C_ACCENT).pack(pady=(0, 10))

        self.chart_canvas = tk.Canvas(sidebar, width=140, height=140, bg=self.C_SIDEBAR, highlightthickness=0)
        self.chart_canvas.pack(pady=10)
        self.update_pie_chart(self.sent_count.get(), self.failed_count.get())

        tk.Frame(sidebar, bg=self.C_BORDER, height=1).pack(fill="x", padx=30, pady=10)
        self.add_side_metric(sidebar, "DELIVERED", self.sent_count, self.C_SUCCESS)
        self.add_side_metric(sidebar, "FAILED", self.failed_count, self.C_STOP)
        self.add_side_metric(sidebar, "REMAINING", self.remain_count, self.C_TEXT_SEC)
        
        tk.Frame(sidebar, bg=self.C_BORDER, height=1).pack(fill="x", padx=30, pady=10)
        time_frame = tk.Frame(sidebar, bg=self.C_SIDEBAR, pady=8)
        time_frame.pack(fill="x", padx=35)
        tk.Label(time_frame, text="EST. TIME", font=("Segoe UI", 8, "bold"), bg=self.C_SIDEBAR, fg=self.C_ACCENT).pack(side="left")
        tk.Label(time_frame, textvariable=self.estimated_time, font=("Segoe UI", 11, "bold"), bg=self.C_SIDEBAR, fg=self.C_TEXT).pack(side="right")

        tk.Frame(sidebar, bg=self.C_BORDER, height=1).pack(fill="x", padx=30, pady=10)
        btn_settings = self.create_side_btn(sidebar, "⚙️ SETTINGS", self.open_settings)
        btn_settings.pack(fill="x", padx=25, pady=5)
        self.bind_hover(btn_settings, self.C_SIDEBAR, self.C_BORDER)

        btn_db = self.create_side_btn(sidebar, "📂 UPLOAD DATABASE", self.select_file)
        btn_db.pack(fill="x", padx=25, pady=5)
        self.bind_hover(btn_db, self.C_SIDEBAR, self.C_BORDER)

        sig_box = tk.Frame(sidebar, bg=self.C_SIDEBAR)
        sig_box.pack(side="bottom", pady=20)
        
        tk.Label(sig_box, text=f"Status: {self.user_plan}", font=("Segoe UI", 12, "bold"), bg=self.C_SIDEBAR, fg=self.C_GOLD).pack(pady=(0, 15))

        tk.Label(sig_box, text="POWERED BY", font=("Segoe UI", 9, "bold"), bg=self.C_SIDEBAR, fg=self.C_TEXT_SEC).pack()
        tk.Label(sig_box, text="KHALED KHEDR", font=("Segoe UI", 16, "bold"), bg=self.C_SIDEBAR, fg=self.C_GOLD).pack()

        main = tk.Frame(self.app_frame, bg=self.C_BG)
        main.pack(side="left", fill="both", expand=True, padx=45, pady=40)

        top_h = tk.Frame(main, bg=self.C_BG); top_h.pack(fill="x", pady=(0, 10))
        tk.Label(top_h, text="Mission Dashboard", font=("Segoe UI", 26, "bold"), bg=self.C_BG, fg=self.C_TEXT).pack(side="left")
        
        action_bar = tk.Frame(main, bg=self.C_CARD, highlightthickness=1, highlightbackground=self.C_ACCENT)
        action_bar.pack(fill="x", pady=(0, 20), ipady=8)
        tk.Label(action_bar, text="⚡ LIVE ACTION ⚡", font=("Segoe UI", 9, "bold"), bg=self.C_CARD, fg=self.C_GOLD).pack(pady=(5,0))
        self.live_action_lbl = tk.Label(action_bar, textvariable=self.file_status, font=("Segoe UI", 16, "bold"), bg=self.C_CARD, fg=self.C_ACCENT)
        self.live_action_lbl.pack(pady=(0,5))

        p_f = tk.Frame(main, bg=self.C_BG); p_f.pack(fill="x", pady=(0, 20))
        style = ttk.Style(); style.theme_use('default')
        style.configure("SiaSend.Horizontal.TProgressbar", thickness=10, background=self.C_ACCENT, troughcolor=self.C_CARD, borderwidth=0)
        style.configure("Vertical.TScrollbar", background=self.C_BORDER, troughcolor=self.C_CARD, bordercolor=self.C_BG, arrowcolor=self.C_TEXT)
        
        self.progress = ttk.Progressbar(p_f, orient="horizontal", mode="determinate", style="SiaSend.Horizontal.TProgressbar")
        self.progress.pack(side="left", fill="x", expand=True)
        tk.Label(p_f, textvariable=self.progress_pct, font=("Segoe UI", 10, "bold"), bg=self.C_BG, fg=self.C_ACCENT).pack(side="right", padx=(15, 0))

        body = tk.Frame(main, bg=self.C_BG); body.pack(fill="both", expand=True); body.columnconfigure(0, weight=3); body.columnconfigure(1, weight=2)

        log_f = tk.LabelFrame(body, text=" LIVE MISSION TRACKER ", bg=self.C_CARD, fg=self.C_TEXT_SEC, font=("Segoe UI", 10, "bold"), padx=15, pady=15, relief="flat", highlightthickness=1, highlightbackground=self.C_BORDER); log_f.grid(row=0, column=0, sticky="nsew", padx=(0, 20))
        
        style.configure("Custom.Treeview", background=self.C_TABLE_BG, foreground=self.C_TABLE_FG, fieldbackground=self.C_TABLE_BG, borderwidth=0, font=("Consolas", 10))
        style.map("Custom.Treeview", background=[("selected", self.C_ACCENT)])
        style.configure("Custom.Treeview.Heading", background=self.C_CARD, foreground=self.C_TEXT, font=("Segoe UI", 9, "bold"), borderwidth=0)

        columns = ("ID", "Name", "Phone", "Wait", "Status")
        self.log_table = ttk.Treeview(log_f, columns=columns, show="headings", style="Custom.Treeview")
        
        self.log_table.heading("ID", text="#")
        self.log_table.heading("Name", text="Contact Name")
        self.log_table.heading("Phone", text="Phone Number")
        self.log_table.heading("Wait", text="Delay")
        self.log_table.heading("Status", text="Status")
        
        self.log_table.column("ID", width=40, anchor="center")
        self.log_table.column("Name", width=150, anchor="w")
        self.log_table.column("Phone", width=120, anchor="center")
        self.log_table.column("Wait", width=60, anchor="center")
        self.log_table.column("Status", width=140, anchor="center") 
        
        table_scroll = AutoScrollbar(log_f, orient="vertical", command=self.log_table.yview)
        self.log_table.configure(yscrollcommand=table_scroll.set)
        self.log_table.pack(side="left", fill="both", expand=True)
        
        self.log_table.tag_configure('SUCCESS', foreground=self.C_SUCCESS)
        self.log_table.tag_configure('ERROR', foreground=self.C_STOP)
        self.log_table.tag_configure('PENDING', foreground=self.C_TEXT_SEC)
        self.log_table.tag_configure('ACTIVE', background=self.C_TABLE_ACTIVE, foreground=self.C_TEXT)

        slot_ui = tk.LabelFrame(body, text=" MESSAGE SLOTS ", bg=self.C_CARD, fg=self.C_TEXT_SEC, font=("Segoe UI", 10, "bold"), padx=15, pady=15, relief="flat", highlightthickness=1, highlightbackground=self.C_BORDER); slot_ui.grid(row=0, column=1, sticky="nsew", padx=(20, 0))
        self.canvas = tk.Canvas(slot_ui, bg=self.C_CARD, highlightthickness=0)
        
        self.scrollbar = AutoScrollbar(slot_ui, orient="vertical", command=self.canvas.yview) 
        self.slots_box = tk.Frame(self.canvas, bg=self.C_CARD)
        self.canvas.create_window((0, 0), window=self.slots_box, anchor="nw", width=380)
        self.canvas.configure(yscrollcommand=self.scrollbar.set); self.canvas.pack(side="left", fill="both", expand=True)
        
        slot_ctrl = tk.Frame(slot_ui, bg=self.C_CARD, pady=10); slot_ctrl.pack(fill="x", side="bottom")
        btn_add = tk.Button(slot_ctrl, text="+ ADD NEW VARIANT", command=self.add_slot, bg=self.C_ACCENT, fg="#FFFFFF", font=("Segoe UI", 9, "bold"), relief="flat", pady=10)
        btn_add.pack(fill="x")
        self.bind_hover(btn_add, self.C_ACCENT, "#008BDB")
        
        btn_save = tk.Button(slot_ctrl, text="SAVE TEMPLATES", command=self.save_config, bg=self.C_SIDEBAR, fg=self.C_TEXT_SEC, font=("Segoe UI", 9, "bold"), relief="flat", pady=8)
        btn_save.pack(fill="x", pady=(5,0))
        self.bind_hover(btn_save, self.C_SIDEBAR, self.C_BORDER)

        tools_frame = tk.Frame(main, bg=self.C_BG)
        tools_frame.pack(fill="x", pady=(20, 0))

        self.preview_f = tk.LabelFrame(tools_frame, text=" 💬 LIVE MESSAGE PREVIEW ", bg=self.C_CARD, fg=self.C_ACCENT, font=("Segoe UI", 10, "bold"), padx=20, pady=15, relief="flat", highlightthickness=1, highlightbackground=self.C_BORDER)
        self.preview_f.pack(fill="x") 
        
        self.lbl_preview_target = tk.Label(self.preview_f, text="🎯 Target: ---", bg=self.C_CARD, fg=self.C_TEXT_SEC, font=("Segoe UI", 10, "bold"))
        self.lbl_preview_target.pack(anchor="w", pady=(0, 8))
        
        self.txt_preview_msg = tk.Text(self.preview_f, height=3, bg=self.C_TABLE_BG, fg=self.C_TEXT, font=("Segoe UI", 11), relief="flat", state="disabled", highlightthickness=1, highlightbackground=self.C_BORDER, padx=10, pady=10)
        self.txt_preview_msg.pack(fill="x")

        act_bar = tk.Frame(main, bg=self.C_BG, pady=20); act_bar.pack(fill="x", side="bottom")
        self.btn_run = tk.Button(act_bar, text="LAUNCH MISSION", bg=self.C_ACCENT, fg="#FFFFFF", font=("Segoe UI", 13, "bold"), relief="flat", width=20, pady=15, command=lambda: self.start_thread("full"))
        self.btn_run.pack(side="left", padx=10)
        self.bind_hover(self.btn_run, self.C_ACCENT, "#008BDB")

        self.btn_pause = tk.Button(act_bar, text="PAUSE", bg=self.C_PAUSE, fg="#FFFFFF", font=("Segoe UI", 11, "bold"), relief="flat", width=12, pady=15, state="disabled", command=self.toggle_pause)
        self.btn_pause.pack(side="left", padx=10)
        
        self.btn_test = tk.Button(act_bar, text="TEST ONE", bg=self.C_CARD, fg=self.C_TEXT, font=("Segoe UI", 11, "bold"), relief="flat", width=12, pady=15, highlightthickness=1, highlightbackground=self.C_TEXT, command=lambda: self.start_thread("test"))
        self.btn_test.pack(side="left", padx=10)
        self.bind_hover(self.btn_test, self.C_CARD, self.C_BORDER)

        self.btn_stop = tk.Button(act_bar, text="STOP MISSION", bg=self.C_STOP, fg="#FFFFFF", font=("Segoe UI", 12, "bold"), relief="flat", width=15, pady=15, state="disabled", command=self.stop_process)
        self.btn_stop.pack(side="right", padx=10)
        self.bind_hover(self.btn_stop, self.C_STOP, "#B82A2A")
        
        self.enable_scroll_magic()

    def _gui_update_preview(self, target_info, message_text):
        self.lbl_preview_target.config(text=f"🎯 Target: {target_info}")
        self.txt_preview_msg.config(state="normal")
        self.txt_preview_msg.delete("1.0", tk.END)
        self.txt_preview_msg.insert("1.0", message_text)
        self.txt_preview_msg.config(state="disabled")

    def enable_scroll_magic(self):
        def _on_mousewheel(event):
            widget = self.root.winfo_containing(event.x_root, event.y_root)
            if not widget: return
            delta = int(-1*(event.delta/120))
            while widget:
                if isinstance(widget, tk.Text):
                    if widget == self.root.focus_get() or widget == self.txt_preview_msg: return 
                elif isinstance(widget, ttk.Treeview):
                    if widget.yview() != (0.0, 1.0): 
                        widget.yview_scroll(delta, "units")
                        return "break"
                elif isinstance(widget, tk.Canvas):
                    if widget.yview() != (0.0, 1.0):
                        widget.yview_scroll(delta, "units")
                        return "break"
                parent_name = widget.winfo_parent()
                if not parent_name: break
                widget = widget._nametowidget(parent_name)
        self.root.bind_all("<MouseWheel>", _on_mousewheel)

    def update_pie_chart(self, s, f):
        self.chart_canvas.delete("all"); total = s + f
        if total == 0: 
            self.chart_canvas.create_oval(10, 10, 130, 130, outline=self.C_BORDER, width=6)
            self.chart_canvas.create_text(70, 70, text="Ready", fill=self.C_TEXT_SEC, font=("Segoe UI", 8, "bold"))
        else:
            s_deg = (s / total) * 360; f_deg = (f / total) * 360
            self.chart_canvas.create_arc(10, 10, 130, 130, start=90, extent=-s_deg, fill=self.C_SUCCESS, outline="")
            self.chart_canvas.create_arc(10, 10, 130, 130, start=90-s_deg, extent=-f_deg, fill=self.C_STOP, outline="")
            self.chart_canvas.create_oval(40, 40, 100, 100, fill=self.C_SIDEBAR, outline="")
            self.chart_canvas.create_text(70, 70, text=f"{int((s/total)*100)}%", fill=self.C_TEXT, font=("Segoe UI", 11, "bold"))

    def create_side_btn(self, p, t, c): return tk.Button(p, text=t, bg=self.C_SIDEBAR, fg=self.C_TEXT, font=("Segoe UI", 9, "bold"), relief="flat", pady=15, anchor="w", padx=25, command=c)
    
    def add_side_metric(self, p, l, v, c):
        f = tk.Frame(p, bg=self.C_SIDEBAR, pady=4); f.pack(fill="x", padx=35)
        tk.Label(f, text=l, font=("Segoe UI", 8, "bold"), bg=self.C_SIDEBAR, fg=self.C_TEXT_SEC).pack(side="left")
        tk.Label(f, textvariable=v, font=("Segoe UI", 11, "bold"), bg=self.C_SIDEBAR, fg=c).pack(side="right")

    def add_slot(self, text=""):
        f = tk.Frame(self.slots_box, bg=self.C_CARD, pady=12); f.pack(fill="x", expand=True); f.columnconfigure(1, weight=1)
        idx = len(self.template_slots) + 1
        tk.Label(f, text=f"#{idx:02}", font=("Segoe UI", 10, "bold"), bg=self.C_CARD, fg=self.C_ACCENT).grid(row=0, column=0, padx=10)
        e = tk.Text(f, font=("Segoe UI", 11), bg=self.C_TABLE_BG, fg=self.C_TEXT, height=3, relief="flat", highlightthickness=1, highlightbackground=self.C_BORDER, insertbackground=self.C_TEXT, padx=10, pady=5, undo=True)
        e.insert("1.0", text); e.grid(row=0, column=1, sticky="nsew")
        self.master_paste_fix(e)
        e.bind("<FocusIn>", lambda event, tw=e: self.set_focused_text(tw))
        tk.Button(f, text="✕", fg="#FFFFFF", bg=self.C_STOP, relief="flat", font=("Segoe UI", 9, "bold"), width=4, command=lambda: self.remove_slot(f, e)).grid(row=0, column=2, padx=10)
        self.template_slots.append((e, f)); self.update_scroll_region()

    def remove_slot(self, frame, entry):
        if len(self.template_slots) <= 1: return
        frame.destroy(); self.template_slots.remove((entry, frame)); self.update_scroll_region()

    def update_scroll_region(self): self.slots_box.update_idletasks(); self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def calculate_eta(self, remaining_count):
        if remaining_count <= 0: return "00:00:00"
        avg_wait = (self.config["min_delay"] + self.config["max_delay"]) / 2
        avg_typing = 10 if self.config.get("anti_ban", True) else 3
        avg_time_per_msg = avg_wait + avg_typing + 12
        return str(datetime.timedelta(seconds=int(remaining_count * avg_time_per_msg)))

    def select_file(self):
        f = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.csv")])
        if f:
            df = pd.read_excel(f) if not f.endswith('.csv') else pd.read_csv(f)
            df['Phone_Clean'] = df['Phone'].astype(str).str.replace(r'\D+', '', regex=True)
            df = df[df['Phone_Clean'].str.len() >= 8]
            
            total_initial = len(df)
            df_exact_dropped = df.drop_duplicates()
            exact_duplicates_count = total_initial - len(df_exact_dropped)
            df_all_dropped = df_exact_dropped.drop_duplicates(subset=['Phone_Clean'])
            partial_duplicates_count = len(df_exact_dropped) - len(df_all_dropped)
            
            if exact_duplicates_count > 0 or partial_duplicates_count > 0:
                dialog = tk.Toplevel(self.root)
                dialog.title("Duplicate Manager")
                dialog.geometry("580x370")
                dialog.configure(bg=self.C_BG)
                dialog.resizable(False, False)
                dialog.grab_set() 
                dialog.update_idletasks()
                x = (dialog.winfo_screenwidth() // 2) - (580 // 2)
                y = (dialog.winfo_screenheight() // 2) - (370 // 2)
                dialog.geometry(f"+{x}+{y}")
                
                tk.Label(dialog, text="⚠️ Warning: Duplicate numbers found in database", font=("Segoe UI", 12, "bold"), bg=self.C_BG, fg=self.C_ACCENT).pack(pady=(20, 10))
                info_text = f"Total Duplicates: {exact_duplicates_count + partial_duplicates_count}\n\n🔹 {exact_duplicates_count} Exact matches.\n🔸 {partial_duplicates_count} Partial matches (Same phone, different name)."
                tk.Label(dialog, text=info_text, font=("Segoe UI", 10, "bold"), justify="center", bg=self.C_BG, fg=self.C_TEXT).pack(pady=5)
                
                choice_var = tk.StringVar(value="all")
                rb_frame = tk.Frame(dialog, bg=self.C_BG)
                rb_frame.pack(fill="x", padx=30, pady=10)
                
                tk.Radiobutton(rb_frame, text="Remove all duplicates (Keep one - Recommended)", variable=choice_var, value="all", bg=self.C_BG, fg=self.C_TEXT, selectcolor=self.C_CARD, activebackground=self.C_BG, activeforeground=self.C_ACCENT, font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=5)
                tk.Radiobutton(rb_frame, text="Remove exact matches only", variable=choice_var, value="exact", bg=self.C_BG, fg=self.C_TEXT, selectcolor=self.C_CARD, activebackground=self.C_BG, activeforeground=self.C_ACCENT, font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=5)
                tk.Radiobutton(rb_frame, text="Don't remove (Send everything)", variable=choice_var, value="none", bg=self.C_BG, fg=self.C_TEXT, selectcolor=self.C_CARD, activebackground=self.C_BG, activeforeground=self.C_ACCENT, font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=5)
                
                def apply_choice():
                    nonlocal df
                    choice = choice_var.get()
                    if choice == "all": df = df_all_dropped.copy()
                    elif choice == "exact": df = df_exact_dropped.copy()
                    dialog.destroy()
                    
                btn_apply = tk.Button(dialog, text="Apply Settings", bg=self.C_ACCENT, fg="#FFFFFF", font=("Segoe UI", 11, "bold"), relief="flat", command=apply_choice, width=20, pady=5)
                btn_apply.pack(pady=10)
                self.bind_hover(btn_apply, self.C_ACCENT, "#008BDB")
                self.root.wait_window(dialog)
            
            df['Phone'] = df['Phone_Clean']
            df.drop(columns=['Phone_Clean'], inplace=True)
            df.reset_index(drop=True, inplace=True)
            
            self.config["last_file"] = f; self.save_config(); self.file_status.set("DB Connected")
            for item in self.log_table.get_children(): self.log_table.delete(item)
            df['Delay_Sec'] = [int(random.uniform(self.config["min_delay"], self.config["max_delay"])) for _ in range(len(df))]
            for index, row in df.iterrows():
                contact_name = str(row.get('Name', row.get('الاسم', row.iloc[0] if len(row) > 0 else "N/A")))
                self.log_table.insert('', 'end', values=(index + 1, contact_name, row['Phone'], f"{row['Delay_Sec']}s", "Pending"), tags=('PENDING',))
            
            self.remain_count.set(len(df)); self.progress["maximum"] = len(df)
            self.estimated_time.set(self.calculate_eta(len(df)))
            self.current_df = df
            
            if not df.empty and self.template_slots:
                first_row = df.iloc[0]
                first_template = self.template_slots[0][0].get("1.0", tk.END).strip()
                name = str(first_row.get('Name', first_row.get('الاسم', '')))
                supp = str(first_row.get('Supplier', first_row.get('المورد', '')))
                preview_msg = first_template.replace("[Name]", name).replace("[الاسم]", name).replace("[Supplier]", supp).replace("[المورد]", supp)
                for col in df.columns:
                    if col != 'Delay_Sec':
                        val = "" if pd.isna(first_row.get(col)) else str(first_row.get(col))
                        preview_msg = preview_msg.replace(f"[{col}]", val)
                phone = "".join(filter(str.isdigit, str(first_row.get('Phone', ''))))
                self._gui_update_preview(f"{name} - {phone}", preview_msg)

    def toggle_pause(self):
        self.is_paused = not self.is_paused
        self.btn_pause.config(text="RESUME" if self.is_paused else "PAUSE", bg=self.C_SUCCESS if self.is_paused else self.C_PAUSE)

    def refresh_ui_for_theme(self):
        old_slots = [e.get("1.0", tk.END).strip() for e, f in self.template_slots if e.get("1.0", tk.END).strip()]
        for widget in self.root.winfo_children():
            if widget == getattr(self, 'settings_win', None): continue
            widget.destroy()
        self.apply_theme_colors()
        self.root.configure(bg=self.C_BG)
        self.setup_ui()
        self.template_slots = []
        for t in old_slots:
            if t: self.add_slot(t)
        if not self.template_slots: self.add_slot()
        if hasattr(self, 'current_df') and not self.current_df.empty:
            for index, row in self.current_df.iterrows():
                contact_name = str(row.get('Name', row.get('الاسم', row.iloc[0] if len(row) > 0 else "N/A")))
                phone = "".join(filter(str.isdigit, str(row.get('Phone', ''))))
                self.log_table.insert('', 'end', values=(index + 1, contact_name, phone, f"{row['Delay_Sec']}s", "Pending"), tags=('PENDING',))

    def open_settings(self):
        if hasattr(self, 'settings_win') and self.settings_win and self.settings_win.winfo_exists():
            self.settings_win.lift()
            self.settings_win.focus_force()
            return
        self.settings_win = tk.Toplevel(self.root)
        self.settings_win.title("Settings")
        screen_h = self.root.winfo_screenheight()
        set_h = min(800, screen_h - 60)
        self.settings_win.geometry(f"550x{set_h}")
        self.settings_win.configure(bg=self.C_BG)

        self.set_canvas = tk.Canvas(self.settings_win, bg=self.C_BG, highlightthickness=0)
        self.set_scroll = AutoScrollbar(self.settings_win, orient="vertical", command=self.set_canvas.yview)
        self.set_canvas.configure(yscrollcommand=self.set_scroll.set)
        self.set_scroll.pack(side="right", fill="y")
        self.set_canvas.pack(side="left", fill="both", expand=True)

        self.set_frame = tk.Frame(self.set_canvas, bg=self.C_BG)
        self.set_window = self.set_canvas.create_window((0, 0), window=self.set_frame, anchor="nw")

        def on_set_frame_config(e): self.set_canvas.configure(scrollregion=self.set_canvas.bbox("all"))
        def on_set_canvas_config(e):
            self.set_canvas.itemconfig(self.set_window, width=e.width)
            if self.set_frame.winfo_reqheight() < e.height: self.set_canvas.itemconfig(self.set_window, height=e.height)
            else: self.set_canvas.itemconfig(self.set_window, height=self.set_frame.winfo_reqheight())

        self.set_frame.bind("<Configure>", on_set_frame_config)
        self.set_canvas.bind("<Configure>", on_set_canvas_config)
        
        tk.Label(self.set_frame, text="CONFIGURATION", font=("Segoe UI", 12, "bold"), bg=self.C_BG, fg=self.C_ACCENT).pack(pady=(20, 5))

        def add_setting_row(parent, label_text, tooltip_text):
            f = tk.Frame(parent, bg=self.C_BG); f.pack(fill="x", padx=10, pady=5) 
            tk.Label(f, text=label_text, bg=self.C_BG, fg=self.C_TEXT).pack(side="left")
            lbl_help = tk.Label(f, text="[?]", fg=self.C_ACCENT, bg=self.C_BG, font=("Segoe UI", 10, "bold"), cursor="hand2")
            lbl_help.pack(side="left", padx=5)
            ToolTip(lbl_help, tooltip_text)
            return f

        sending_frame = tk.LabelFrame(self.set_frame, text=" Sending Configuration ", bg=self.C_BG, fg=self.C_TEXT_SEC, font=("Segoe UI", 9, "bold"), padx=10, pady=5)
        sending_frame.pack(fill="x", padx=40, pady=(0, 10))
        f1 = add_setting_row(sending_frame, "Min Delay (sec):", "Minimum wait time between messages.")
        e_min = tk.Entry(f1, width=10); e_min.insert(0, str(self.config.get("min_delay", 20))); e_min.pack(side="right")
        f2 = add_setting_row(sending_frame, "Max Delay (sec):", "Maximum random wait time to prevent bans.")
        e_max = tk.Entry(f2, width=10); e_max.insert(0, str(self.config.get("max_delay", 35))); e_max.pack(side="right")
        f_cc = add_setting_row(sending_frame, "Default Country Code:", "Prefix added to numbers without a country code.")
        e_cc = tk.Entry(f_cc, width=10); e_cc.insert(0, str(self.config.get("country_code", "20"))); e_cc.pack(side="right")

        protection_frame = tk.LabelFrame(self.set_frame, text=" Protection & Anti-Ban ", bg=self.C_BG, fg=self.C_TEXT_SEC, font=("Segoe UI", 9, "bold"), padx=10, pady=5)
        protection_frame.pack(fill="x", padx=40, pady=(0, 10))
        f3 = add_setting_row(protection_frame, "Anti-Ban Mode:", "Types message character by character to mimic human behavior.")
        self.anti_ban_var = tk.BooleanVar(value=self.config.get("anti_ban", True))
        tk.Checkbutton(f3, text="Enable", variable=self.anti_ban_var, bg=self.C_BG, fg=self.C_TEXT, selectcolor=self.C_CARD, activebackground=self.C_BG).pack(side="right")

        f_sleep = tk.Frame(protection_frame, bg=self.C_BG); f_sleep.pack(fill="x", padx=10, pady=5)
        sleep_top = tk.Frame(f_sleep, bg=self.C_BG); sleep_top.pack(fill="x")
        lbl_sleep_help = tk.Label(sleep_top, text="[?]", fg=self.C_ACCENT, bg=self.C_BG, font=("Segoe UI", 10, "bold"), cursor="hand2")
        lbl_sleep_help.pack(side="left", padx=5)
        ToolTip(lbl_sleep_help, "Smart Sleep: Bot pauses completely after a set number of messages.")
        self.sleep_var = tk.BooleanVar(value=self.config.get("smart_sleep", False))
        cb_sleep = tk.Checkbutton(sleep_top, text="Smart Sleep Enable", variable=self.sleep_var, bg=self.C_BG, fg=self.C_TEXT, selectcolor=self.C_CARD, activebackground=self.C_BG)
        cb_sleep.pack(side="left", padx=5)

        sleep_bot = tk.Frame(f_sleep, bg=self.C_BG); sleep_bot.pack(fill="x", pady=5, padx=20)
        tk.Label(sleep_bot, text="Pause for", bg=self.C_BG, fg=self.C_TEXT).pack(side="left")
        e_sleep_min = tk.Entry(sleep_bot, width=4); e_sleep_min.insert(0, str(self.config.get("sleep_mins", 5))); e_sleep_min.pack(side="left", padx=5)
        tk.Label(sleep_bot, text="mins after", bg=self.C_BG, fg=self.C_TEXT).pack(side="left")
        e_sleep_msgs = tk.Entry(sleep_bot, width=4); e_sleep_msgs.insert(0, str(self.config.get("sleep_msgs", 50))); e_sleep_msgs.pack(side="left", padx=5)
        tk.Label(sleep_bot, text="msgs.", bg=self.C_BG, fg=self.C_TEXT).pack(side="left")

        appearance_frame = tk.LabelFrame(self.set_frame, text=" Appearance ", bg=self.C_BG, fg=self.C_TEXT_SEC, font=("Segoe UI", 9, "bold"), padx=10, pady=5)
        appearance_frame.pack(fill="x", padx=40, pady=(0, 10))
        f_theme = add_setting_row(appearance_frame, "UI Theme:", "Switch UI between Light and Dark mode.")
        self.theme_var = tk.StringVar(value=self.config.get("theme", "Dark"))
        tk.Radiobutton(f_theme, text="Light", variable=self.theme_var, value="Light", bg=self.C_BG, fg=self.C_TEXT, selectcolor=self.C_CARD, activebackground=self.C_BG).pack(side="right", padx=10)
        tk.Radiobutton(f_theme, text="Dark", variable=self.theme_var, value="Dark", bg=self.C_BG, fg=self.C_TEXT, selectcolor=self.C_CARD, activebackground=self.C_BG).pack(side="right")

        rotation_frame = tk.LabelFrame(self.set_frame, text=" Multi-Account Rotation ", bg=self.C_BG, fg=self.C_TEXT_SEC, font=("Segoe UI", 9, "bold"), padx=10, pady=5)
        rotation_frame.pack(fill="x", padx=40, pady=(0, 10))
        f_rot = tk.Frame(rotation_frame, bg=self.C_BG); f_rot.pack(fill="x", padx=10, pady=5)
        rot_top = tk.Frame(f_rot, bg=self.C_BG); rot_top.pack(fill="x")
        lbl_rot_help = tk.Label(rot_top, text="[?]", fg=self.C_ACCENT, bg=self.C_BG, font=("Segoe UI", 10, "bold"), cursor="hand2")
        lbl_rot_help.pack(side="left", padx=5)
        ToolTip(lbl_rot_help, "Switches between multiple WhatsApp accounts automatically to spread sending volume.")
        self.rot_var = tk.BooleanVar(value=self.config.get("multi_account", False))
        cb_rot = tk.Checkbutton(rot_top, text="Enable Rotation", variable=self.rot_var, bg=self.C_BG, fg=self.C_TEXT, selectcolor=self.C_CARD, activebackground=self.C_BG)
        cb_rot.pack(side="left", padx=5)

        rot_bot = tk.Frame(f_rot, bg=self.C_BG); rot_bot.pack(fill="x", pady=5, padx=20)
        tk.Label(rot_bot, text="Rotate between", bg=self.C_BG, fg=self.C_TEXT).pack(side="left")
        e_num_accs = tk.Entry(rot_bot, width=4); e_num_accs.insert(0, str(self.config.get("num_accounts", 2))); e_num_accs.pack(side="left", padx=5)
        tk.Label(rot_bot, text="accounts, every", bg=self.C_BG, fg=self.C_TEXT).pack(side="left")
        e_rot_msgs = tk.Entry(rot_bot, width=4); e_rot_msgs.insert(0, str(self.config.get("rotate_after", 20))); e_rot_msgs.pack(side="left", padx=5)
        tk.Label(rot_bot, text="msgs.", bg=self.C_BG, fg=self.C_TEXT).pack(side="left")

        def apply_settings():
            old_theme = self.config.get("theme")
            self.config.update({
                "min_delay": int(e_min.get()), "max_delay": int(e_max.get()), "anti_ban": self.anti_ban_var.get(), "country_code": e_cc.get().strip().replace("+", ""),
                "smart_sleep": self.sleep_var.get(), "sleep_mins": int(e_sleep_min.get()), "sleep_msgs": int(e_sleep_msgs.get()), "theme": self.theme_var.get(),
                "multi_account": self.rot_var.get(), "num_accounts": int(e_num_accs.get()), "rotate_after": int(e_rot_msgs.get())
            })
            self.save_config()
            if hasattr(self, 'current_df'):
                self.current_df['Delay_Sec'] = [int(random.uniform(self.config["min_delay"], self.config["max_delay"])) for _ in range(len(self.current_df))]
                for i, item in enumerate(self.log_table.get_children()):
                    row_vals = list(self.log_table.item(item, 'values'))
                    row_vals[3] = f"{self.current_df.iloc[i]['Delay_Sec']}s"
                    self.log_table.item(item, values=row_vals)
                self.estimated_time.set(self.calculate_eta(self.remain_count.get()))
            if old_theme != self.theme_var.get(): self.refresh_ui_for_theme() 
            self.settings_win.destroy()

        btn_apply = tk.Button(self.set_frame, text="APPLY", bg=self.C_ACCENT, fg="#FFFFFF", relief="flat", width=20, pady=10, command=apply_settings)
        btn_apply.pack(pady=(10, 20))
        self.bind_hover(btn_apply, self.C_ACCENT, "#008BDB")

    def load_initial_templates(self):
        self.template_slots = []
        for t in self.config.get("templates", []): self.add_slot(t)
        if not self.template_slots: self.add_slot()

    def stop_process(self): 
        self.stop_requested = True; self.is_paused = False

    def _gui_update_table(self, item, values, tags, see=True):
        self.log_table.item(item, values=values, tags=tags)
        if see: self.log_table.see(item)

    def _gui_update_progress(self, val, pct, remaining, eta):
        self.progress["value"] = val
        self.progress_pct.set(pct)
        self.remain_count.set(remaining)
        self.estimated_time.set(eta)

    def _gui_finish_mission(self):
        self.is_running = False
        self.btn_run.config(state="normal", bg=self.C_ACCENT)
        self.btn_stop.config(state="disabled", bg=self.C_SIDEBAR)
        self.btn_pause.config(state="disabled", bg=self.C_SIDEBAR)
        self.estimated_time.set("00:00:00")
        self.file_status.set("System Ready")
        winsound.Beep(1000, 500)
        if messagebox.askyesno("Mission Complete", "Mission Complete! ✅\nDo you want to save a summary report?"):
            self.export_report()

    def _gui_error_finish(self):
        self.is_running = False
        self.btn_run.config(state="normal", bg=self.C_ACCENT)
        self.btn_stop.config(state="disabled", bg=self.C_SIDEBAR)
        self.btn_pause.config(state="disabled", bg=self.C_SIDEBAR)
        self.estimated_time.set("00:00:00")
        self.file_status.set("System Error")

    def start_thread(self, mode):
        if not self.is_running:
            if not hasattr(self, 'current_df') or self.current_df.empty: 
                messagebox.showerror("Error", "Please upload a database first!"); return
            
            self.is_running = True; self.stop_requested = False; self.is_paused = False
            self.mission_report = [] 
            self.sent_count.set(0); self.failed_count.set(0); self.update_pie_chart(0, 0)
            self.btn_run.config(state="disabled", bg=self.C_SIDEBAR); self.btn_stop.config(state="normal", bg=self.C_STOP); self.btn_pause.config(state="normal", bg=self.C_PAUSE)
            
            temps = [e.get("1.0", tk.END).strip() for e, f in self.template_slots if e.get("1.0", tk.END).strip()]
            items = self.log_table.get_children()
            
            for item in items:
                v = list(self.log_table.item(item, 'values'))
                self.log_table.item(item, values=(v[0], v[1], v[2], v[3], "Pending"), tags=('PENDING',))
            
            threading.Thread(target=self.run_bot, args=(mode, temps, items), daemon=True).start()

    def run_bot(self, mode, temps, items):
        driver = None
        current_account_index = 1
        msgs_on_current_account = 0

        def human_click(current_driver, element):
            try:
                ActionChains(current_driver).move_to_element(element).pause(random.uniform(0.1, 0.4)).click().perform()
            except Exception:
                element.click()

        def init_whatsapp_driver(acc_index):
            nonlocal driver
            if driver:
                try: driver.quit()
                except Exception: pass
            
            self.root.after(0, lambda: self.file_status.set("🌐 Preparing Chrome Profile..."))
            options = Options()
            profile_path = os.environ.get("LOCALAPPDATA")
            profile_name = f"SiaSendProfile_{acc_index}" if self.config.get("multi_account") else "SiaSendProfile"
            
            full_profile_path = os.path.join(profile_path, profile_name)
            options.add_argument(f"--user-data-dir={full_profile_path}")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_argument("--start-maximized")
            options.add_argument("--disable-infobars")
            options.add_argument("--disable-extensions")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
            
            d = self.get_driver(options)
            try:
                d.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                    "source": "Object.defineProperty(navigator, 'webdriver', { get: () => undefined });"
                })
            except Exception: pass

            self.root.after(0, lambda: self.file_status.set("⏳ Waiting for WhatsApp Login..."))
            d.get("https://web.whatsapp.com")
            
            login_event = threading.Event()
            def ask_login():
                messagebox.showinfo("Authentication", f"[{profile_name}]\nPlease scan QR code (if needed) and wait for chats to load, then click OK to start sending.")
                login_event.set()
            self.root.after(0, ask_login)
            login_event.wait() 
            
            self.root.after(0, lambda: self.file_status.set("✅ WhatsApp is Ready!"))
            return d

        try:
            self.root.after(0, lambda: self.file_status.set("🚀 Start launching mission..."))
            df_to_process = self.current_df.copy()
            if mode == "test": df_to_process = df_to_process.head(1)
            driver = init_whatsapp_driver(current_account_index)
            
            for i in range(len(df_to_process)):

                row = df_to_process.iloc[i]
                phone = "".join(filter(str.isdigit, str(row.get('Phone', ''))))
                name = str(row.get('Name', row.get('الاسم', '')))
                cc = str(self.config.get("country_code", ""))
                if cc and phone and not phone.startswith(cc) and len(phone) <= 11: phone = f"{cc}{phone}"
                
                wait_time_str = f"{row['Delay_Sec']}s"
                current_item = items[i]
                
                if self.config.get("smart_sleep") and i > 0 and i % self.config.get("sleep_msgs", 50) == 0:
                    sleep_mins = self.config.get("sleep_mins", 5)
                    self.root.after(0, lambda: self.file_status.set(f"💤 Smart Sleep for {sleep_mins} mins..."))
                    self.root.after(0, lambda it=current_item, v=(i+1, name, phone, wait_time_str, f"Sleeping..."): self._gui_update_table(it, v, ('ACTIVE',)))
                    sleep_sec = sleep_mins * 60
                    while sleep_sec > 0:
                        if self.stop_requested: break
                        time.sleep(1); sleep_sec -= 1
                
                if self.config.get("multi_account") and msgs_on_current_account >= self.config.get("rotate_after", 20):
                    msgs_on_current_account = 0
                    current_account_index += 1
                    if current_account_index > self.config.get("num_accounts", 2): current_account_index = 1
                    self.root.after(0, lambda: self.file_status.set("🔄 Switching Account..."))
                    driver = init_whatsapp_driver(current_account_index)
                
                while not self.check_internet():
                    if self.stop_requested: break
                    if not self.is_paused: self.toggle_pause()
                    time.sleep(2)
                
                while self.is_paused:
                    if self.stop_requested: break
                    time.sleep(1)
                if self.stop_requested: break
                
                self.root.after(0, lambda p=phone: self.file_status.set(f"💬 Sending to {p}..."))
                self.root.after(0, lambda it=current_item, v=(i+1, name, phone, wait_time_str, "Sending..."): self._gui_update_table(it, v, ('ACTIVE',)))
                err_category = "Failed" 
                
                try:
                    msg = random.choice(temps)
                    supp = str(row.get('Supplier', row.get('المورد', '')))
                    msg = msg.replace("[Name]", name).replace("[الاسم]", name).replace("[Supplier]", supp).replace("[المورد]", supp)
                    
                    for col in df_to_process.columns:
                        if col != 'Delay_Sec':
                            val = "" if pd.isna(row.get(col)) else str(row.get(col))
                            msg = msg.replace(f"[{col}]", val)
                    
                    self.root.after(0, lambda t=f"{name} - {phone}", m=msg: self._gui_update_preview(t, m))
                    driver.get(f"https://web.whatsapp.com/send?phone={phone}")
                    
                    try:
                        xpath = '//div[@contenteditable="true"][@data-tab="10" or @title="Type a message"]'
                        box = WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.XPATH, xpath)))
                    except TimeoutException:
                        try:
                            invalid_btn = driver.find_element(By.XPATH, '//div[@role="button"]//span[contains(text(), "OK") or contains(text(), "موافق")]')
                            human_click(driver, invalid_btn)
                            err_category = "Invalid Number"
                            raise Exception("Invalid Number")
                        except Exception:
                            err_category = "Timeout"
                            raise Exception("Timeout")

                    if self.config.get("anti_ban", True):
                        time.sleep(random.uniform(1, 1.5))
                        for char in msg:
                            if char == '\n': box.send_keys(Keys.SHIFT + Keys.ENTER)
                            else: box.send_keys(char)
                            time.sleep(random.uniform(0.005, 0.03)) 
                        time.sleep(0.5); box.send_keys(Keys.ENTER)
                    else:
                        driver.execute_script("arguments[0].innerHTML = '';", box)
                        box.send_keys(msg); time.sleep(1); box.send_keys(Keys.ENTER)
                        
                    time.sleep(1.5)
                    self.sent_count.set(self.sent_count.get() + 1)
                    self.total_lifetime_sent += 1 
                    
                    self.mission_report.append({"Phone": phone, "Status": "Delivered"})
                    self.root.after(0, lambda it=current_item, v=(i+1, name, phone, wait_time_str, "Delivered"): self._gui_update_table(it, v, ('SUCCESS',)))
                    
                except Exception: 
                    self.failed_count.set(self.failed_count.get() + 1)
                    self.mission_report.append({"Phone": phone, "Status": err_category})
                    self.root.after(0, lambda it=current_item, v=(i+1, name, phone, wait_time_str, err_category): self._gui_update_table(it, v, ('ERROR',)))
                
                msgs_on_current_account += 1
                self.root.after(0, lambda s=self.sent_count.get(), f=self.failed_count.get(): self.update_pie_chart(s, f))
                
                val = i + 1; pct = f"{int(((i+1)/len(df_to_process))*100)}%"; rem = len(df_to_process) - (i + 1); eta = self.calculate_eta(rem)
                self.root.after(0, lambda v=val, p=pct, r=rem, e=eta: self._gui_update_progress(v, p, r, e))
                
                if i < len(df_to_process) - 1:
                    time.sleep(row['Delay_Sec'])
                
            self.root.after(0, self._gui_finish_mission)
            
        except Exception as system_error:
            self.root.after(0, lambda: messagebox.showerror("System Error", f"Fatal error:\n\n{str(system_error)}"))
            self.root.after(0, self._gui_error_finish)
                
        finally:
            if driver: driver.quit()

    def export_report(self):
        if not self.mission_report: return
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV file", "*.csv")], title="Save Mission Report")
        if file_path:
            pd.DataFrame(self.mission_report).to_csv(file_path, index=False, encoding='utf-8-sig')

# ==========================================
# --- 3. Single Instance & Launch ---
# ==========================================
_instance_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
try:
    _instance_socket.bind(("127.0.0.1", 65432))
except socket.error:
    tk.Tk().withdraw()
    messagebox.showerror("Instance Error", "SiaSend is already running!")
    sys.exit()

if __name__ == "__main__":
    root = tk.Tk()
    enable_dark_title_bar(root)
    app = SiaSend(root)
    root.mainloop()

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from PIL import Image, ImageTk
import os
import json
import sys
from openpyxl import Workbook
from openpyxl.styles import Font

# --- Nagarkot GUI Constants ---
THEME_COLOR = "#0056b3"  # Nagarkot Blue
BG_COLOR = "#ffffff"     # White
TEXT_COLOR = "#000000"
SECONDARY_BG = "#f0f0f0"
SETTINGS_FILE = "settings.json"

def get_base_path():
    """Returns the base path for READ-ONLY assets (logo, codes) inside the EXE bundle."""
    if hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

def get_asset_path(filename):
    """Returns the path for assets bundled inside the exe."""
    return os.path.join(get_base_path(), filename)

def get_data_dir():
    """Returns the path for WRITABLE persistent data in Windows AppData."""
    data_dir = os.path.join(os.getenv('APPDATA'), "VoucherProcessor")
    if not os.path.exists(data_dir):
        os.makedirs(data_dir, exist_ok=True)
    return data_dir

def get_settings_path():
    """Returns the hidden settings path in AppData."""
    return os.path.join(get_data_dir(), SETTINGS_FILE)

def get_history_path():
    """Returns the hidden history log path in AppData."""
    return os.path.join(get_data_dir(), "download_history.json")

class ToolTip(object):
    def __init__(self, widget, text='widget info'):
        self.waittime = 500     # miliseconds
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.waittime, self.showtip)

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def showtip(self, event=None):
        x = y = 0
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        
        self.tw = tk.Toplevel(self.widget)
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(self.tw, text=self.text, justify='left',
                       background="#ffffe0", relief='solid', borderwidth=1,
                       font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tw
        self.tw = None
        if tw:
            tw.destroy()

class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.canvas = tk.Canvas(self, bg=BG_COLOR, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, style="Card.TFrame")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self._update_scrollregion()
        )

        self.window_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        # Don't pack scrollbar yet - will be shown dynamically
        
        self.bind("<Enter>", self._bind_mousewheel)
        self.bind("<Leave>", self._unbind_mousewheel)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

    def _bind_mousewheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")

    def _update_scrollregion(self):
        """Update scroll region and show/hide scrollbar based on content size"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self._toggle_scrollbar()

    def _toggle_scrollbar(self):
        """Show or hide scrollbar based on whether content exceeds visible area"""
        bbox = self.canvas.bbox("all")
        if bbox:
            content_height = bbox[3] - bbox[1]
            canvas_height = self.canvas.winfo_height()
            
            if content_height > canvas_height:
                # Content exceeds visible area - show scrollbar
                self.scrollbar.pack(side="right", fill="y")
            else:
                # Content fits - hide scrollbar
                self.scrollbar.pack_forget()

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.window_id, width=event.width)
        self._toggle_scrollbar()

    def _on_mousewheel(self, event):
        # Only scroll if content height exceeds canvas height
        bbox = self.canvas.bbox("all")
        if bbox:
            content_height = bbox[3] - bbox[1]
            canvas_height = self.canvas.winfo_height()
            
            # Only allow scrolling if content is taller than visible area
            if content_height > canvas_height:
                self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

class SearchableEntry(ttk.Frame):
    _active_popup = None  # Track the currently open dropdown globally

    def __init__(self, master, values, command=None, width=None, **kwargs):
        super().__init__(master, style="Card.TFrame")
        self.values = sorted(values)
        self.filtered_values = self.values
        self.command = command
        
        self.entry = ttk.Entry(self, **kwargs)
        if width:
            self.entry.config(width=int(width/7))
        # Start in readonly state as requested
        self.entry.config(state="readonly")
        self.entry.pack(fill="x", expand=True)
        
        self.entry.bind("<KeyRelease>", self.on_key_release)
        # suggestions now only show on click or typing, not automatic focus 
        # to prevent re-opening when window is restored
        self.entry.bind("<Button-1>", self.show_suggestions) 
        
        self.popup = None
        self.listbox = None
        self.root_click_id = None
        self.root_unmap_id = None
        self.root_config_id = None
        self._closing = False

    def on_key_release(self, event):
        if event.keysym == "Escape":
            self.close_popup()
            return
            
        if event.keysym == "Down":
            if self.listbox and self.listbox.size() > 0:
                curr = self.listbox.curselection()
                idx = min(curr[0] + 1, self.listbox.size() - 1) if curr else 0
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(idx)
                self.listbox.see(idx)
                self.listbox.activate(idx)
            return

        if event.keysym == "Up":
            if self.listbox and self.listbox.size() > 0:
                curr = self.listbox.curselection()
                idx = max(curr[0] - 1, 0) if curr else 0
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(idx)
                self.listbox.see(idx)
                self.listbox.activate(idx)
            return

        if event.keysym == "Return":
            if self.popup and self.listbox:
                self.on_select(None)
            else:
                val = self.entry.get()
                if self.command:
                    self.command(val)
            return

        typed = self.entry.get().strip().upper()
        if typed == "":
            self.filtered_values = self.values
        else:
            self.filtered_values = [v for v in self.values if typed in v.upper()]
        
        self.show_suggestions()

    def show_suggestions(self, event=None):
        if hasattr(self, '_closing') and self._closing:
            return
            
        # Close any other open dropdown before showing this one
        if SearchableEntry._active_popup and SearchableEntry._active_popup != self:
            try:
                SearchableEntry._active_popup.close_popup()
            except:
                pass
        
        # Temporarily enable to allow typing/interaction
        self.entry.config(state="normal")
        
        if not self.filtered_values:
            # Don't close immediately if active, might be backspacing
            pass

        x = self.entry.winfo_rootx()
        y = self.entry.winfo_rooty() + self.entry.winfo_height()
        width = self.entry.winfo_width()
        
        if not self.popup:
            self.popup = tk.Toplevel(self)
            SearchableEntry._active_popup = self
            self.popup.wm_overrideredirect(True)
            self.popup.attributes("-topmost", True)
            
            bg_color = "#ffffff"
            fg_color = "#000000"
            
            self.list_frame = tk.Frame(self.popup, bg=bg_color, relief="solid", borderwidth=1)
            self.list_frame.pack(fill="both", expand=True)
            
            self.scrollbar = tk.Scrollbar(self.list_frame)
            self.scrollbar.pack(side="right", fill="y")
            
            self.listbox = tk.Listbox(self.list_frame, font=("Arial", 9), bg=bg_color, fg=fg_color, 
                                    selectbackground=THEME_COLOR, selectforeground="white", highlightthickness=0,
                                    yscrollcommand=self.scrollbar.set)
            self.listbox.pack(side="left", fill="both", expand=True)
            self.scrollbar.config(command=self.listbox.yview)
            
            self.listbox.bind("<ButtonRelease-1>", self.on_select)
            self.listbox.bind("<Double-Button-1>", self.on_select)

            root = self.winfo_toplevel()
            self.root_click_id = root.bind("<Button-1>", self.on_root_click, add="+")
            self.root_unmap_id = root.bind("<Unmap>", lambda e: self.close_popup(), add="+")
            self.root_config_id = root.bind("<Configure>", lambda e: self.close_popup(), add="+")

        height = min(len(self.filtered_values) * 20 + 2, 200)
        self.popup.wm_geometry(f"{width}x{height}+{x}+{y}")
        
        self.listbox.delete(0, tk.END)
        for val in self.filtered_values:
            self.listbox.insert(tk.END, val)
        
        if self.filtered_values:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(0)
            self.listbox.activate(0)

    def on_select(self, event):
        if self.listbox and self.listbox.curselection():
            index = self.listbox.curselection()[0]
            selected = self.listbox.get(index)
            self.set(selected)
            if self.command:
                self.command(selected)
        
        # Lock back to readonly after selection
        self.entry.config(state="readonly")
        self.after(50, self.close_popup)

    def on_root_click(self, event):
        if not self.popup or not self.entry.winfo_exists():
            return
            
        x, y = event.x_root, event.y_root
        
        try:
            ex = self.entry.winfo_rootx()
            ey = self.entry.winfo_rooty()
            ew = self.entry.winfo_width()
            eh = self.entry.winfo_height()
            
            if ex <= x <= ex + ew and ey <= y <= ey + eh:
                return

            px = self.popup.winfo_rootx()
            py = self.popup.winfo_rooty()
            pw = self.popup.winfo_width()
            ph = self.popup.winfo_height()
            
            if px <= x <= px + pw and py <= y <= py + ph:
                return
        except tk.TclError:
            pass
            
        # If clicked outside, lock the entry and close dropdown
        self.entry.config(state="readonly")
        self.close_popup()

    def close_popup(self):
        if self.popup:
            self._closing = True
            root = self.winfo_toplevel()
            try:
                if self.root_click_id: root.unbind("<Button-1>", self.root_click_id)
                if self.root_unmap_id: root.unbind("<Unmap>", self.root_unmap_id)
                if self.root_config_id: root.unbind("<Configure>", self.root_config_id)
            except:
                pass
            
            self.root_click_id = None
            self.root_unmap_id = None
            self.root_config_id = None

            try:
                self.popup.destroy()
            except:
                pass
            self.popup = None
            if SearchableEntry._active_popup == self:
                SearchableEntry._active_popup = None
            self.after(300, self._reset_closing)
        
        # Ensure it's readonly when closed
        if self.entry.winfo_exists():
            self.entry.config(state="readonly")

    def _reset_closing(self):
        self._closing = False

    def set(self, value, trigger_command=False):
        # Helper to set text regardless of current state
        orig_state = str(self.entry.cget("state"))
        self.entry.config(state="normal")
        self.entry.delete(0, tk.END)
        self.entry.insert(0, str(value))
        self.entry.config(state=orig_state)
        if trigger_command and self.command:
            self.command(value)

    def get(self):
        return self.entry.get()

class ReimbursementApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Voucher Processor")
        self.geometry("1400x850")
        self.state('zoomed')
        self.configure(bg=BG_COLOR)
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("TFrame", background=BG_COLOR)
        self.style.configure("TLabel", background=BG_COLOR, font=("Arial", 10), foreground=TEXT_COLOR)
        self.style.configure("TButton", font=("Arial", 10))
        self.style.configure("Header.TLabel", font=("Arial", 16, "bold"), foreground=THEME_COLOR, background=BG_COLOR)
        self.style.configure("SubHeader.TLabel", font=("Arial", 10), foreground="#666666", background=BG_COLOR)
        self.style.configure("Card.TFrame", background=BG_COLOR)
        self.style.configure("Row.TFrame", background=BG_COLOR)
        
        self.setup_header()

        self.reimbursement_data = None
        self.expense_details = None
        self.exp_codes = []
        self.selected_codes = {} 
        self.selected_branch = tk.StringVar(value="HO")
        self.selected_cr_dr = tk.StringVar(value="DR")
        self.fiscal_year = tk.StringVar(value="25-26")
        self.load_settings()

        self.load_expense_codes()
        self.current_mode = None 
        
        self.content_frame = ttk.Frame(self)
        self.content_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        self.setup_footer()
        self.setup_selection_ui()
        self.master_code = ""

    def setup_header(self):
        header_frame = ttk.Frame(self, cursor="arrow")
        header_frame.pack(fill="x", padx=20, pady=10)
        
        upper_container = tk.Frame(header_frame, bg=BG_COLOR, height=50)
        upper_container.pack(fill="x")
        upper_container.pack_propagate(False)

        try:
            logo_path = get_asset_path("Nagarkot Logo.png")
            if os.path.exists(logo_path):
                pil_img = Image.open(logo_path)
                h_size = 35
                w_size = int((h_size / float(pil_img.size[1])) * float(pil_img.size[0]))
                pil_img = pil_img.resize((w_size, h_size), Image.Resampling.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(pil_img)
                tk.Label(upper_container, image=self.logo_img, bg=BG_COLOR).pack(side="left", anchor="nw")
            else:
                tk.Label(upper_container, text="[LOGO MISSING]", fg="red", bg=BG_COLOR).pack(side="left")
        except Exception as e:
            print(f"Logo Error: {e}")
        
        title_frame = ttk.Frame(upper_container)
        title_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        ttk.Label(title_frame, text="NAGARKOT FORWARDERS", style="Header.TLabel").pack(anchor="center")
        ttk.Label(title_frame, text="Voucher Automation System", style="SubHeader.TLabel").pack(anchor="center")
        
    def setup_footer(self):
        footer_frame = ttk.Frame(self)
        footer_frame.pack(side="bottom", fill="x", padx=20, pady=10)
        
        ttk.Label(footer_frame, text="© Nagarkot Forwarders Pvt Ltd", font=("Arial", 8), foreground="gray").pack(side="left")

        self.generate_btn = tk.Button(footer_frame, text="Generate Logisys Excel", 
                                      command=self.generate_output, state="disabled",
                                      bg=THEME_COLOR, fg="white", font=("Arial", 10, "bold"), 
                                      relief="flat", padx=15, pady=5)

    def load_settings(self):
        settings_path = get_settings_path()
        if os.path.exists(settings_path):
            try:
                with open(settings_path, "r") as f:
                    settings = json.load(f)
                    if "fiscal_year" in settings:
                        self.fiscal_year.set(settings["fiscal_year"])
            except Exception as e:
                print(f"Error loading settings: {e}")

    def save_settings(self, *args):
        settings = {"fiscal_year": self.fiscal_year.get()}
        settings_path = get_settings_path()
        try:
            with open(settings_path, "w") as f:
                json.dump(settings, f)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def load_history(self):
        """Loads the list of already processed Voucher Numbers."""
        path = get_history_path()
        if os.path.exists(path):
            try:
                with open(path, "r") as f:
                    return json.load(f)
            except:
                return {}
        return {}

    def save_to_history(self, trans_id, mode):
        """Saves a processed Voucher Number to the local history log."""
        path = get_history_path()
        history = self.load_history()
        # Key by ID, store mode and timestamp
        history[str(trans_id)] = {
            "mode": mode,
            "timestamp": pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        try:
            with open(path, "w") as f:
                json.dump(history, f, indent=4)
        except Exception as e:
            print(f"Error saving history: {e}")

    def refresh_previews(self, *args):
        if not hasattr(self, 'row_ui_refs') or not self.row_ui_refs:
            return
        for sub_id, refs in self.row_ui_refs.items():
            if not refs.get('job_val'):
                continue
            lob_val = self.lob_selection[sub_id].get() if sub_id in self.lob_selection else "GEN"
            new_cc, new_job_fmt = self.get_lob_details(lob_val, refs['branch'], refs['job_val'], self.current_mode)
            if refs.get('cc_label'): refs['cc_label'].configure(text=f"CC: {new_cc}")
            if refs.get('badge_label'): refs['badge_label'].configure(text=f"[{new_job_fmt}]")

    def load_expense_codes(self):
        excel_path = get_asset_path("Logisys _ Indirect Exp Codes.xlsx")
        try:
            if os.path.exists(excel_path):
                df_codes = pd.read_excel(excel_path)
                if 'Particulars' in df_codes.columns:
                    self.exp_codes = df_codes['Particulars'].dropna().astype(str).unique().tolist()
                else:
                    self.exp_codes = df_codes.iloc[:, 0].dropna().astype(str).unique().tolist()
            else:
                raise FileNotFoundError
        except Exception as e:
            self.exp_codes = [
                "ADVERTISEMENT EXPENSES", "BUSINESS PROMOTION EXPENSES (GST)", "CONVEYANCE EXPENSES",
                "ELECTRICITY EXPENSES", "GENERAL EXPENSES", "INTERNET EXPENSES",
                "OFFICE MAINTENANCE EXPENSES", "PRINTING & STATIONERY EXPENSES", "RENT EXPENSES",
                "REPAIR & MAINTENANCE EXPENSES", "STAFF WELFARE EXPENSES", "TELEPHONE EXPENSES",
                "TRAVELLING EXPENSES", "WATER CHARGES"
            ]

        self.warehouse_codes = [
            "3PL _ F&B Client Expenses",
            "3PL _ F&B Overtime Expenses",
            "3PL _ F&B Tea Expenses",
            "3PL _ F&B Water Expenses",
            "3PL _ Housekeeping Expenses",
            "3PL _ Manpower Other Expenses",
            "3PL _ Other Misc. Expenses",
            "3PL _ Repairs Expenses",
            "3PL _ Stationery Expenses",
            "3PL Manpower Overtime Expenses",
            "Diesel For Genset Exp",
            "Labour Supply Service Exp"
        ]
        
        self.voucher_codes = ["C&F CHARGES (E)"]

    def setup_selection_ui(self):
        for widget in self.content_frame.winfo_children():
            widget.destroy()
            
        self.generate_btn.pack_forget()

        container = ttk.Frame(self.content_frame)
        container.place(relx=0.5, rely=0.5, anchor="center")

        ttk.Label(container, text="Choose Workflow", font=("Arial", 14, "bold"), foreground="#444").pack(pady=(0, 20))

        card_container = ttk.Frame(container)
        card_container.pack()

        reimb_card = tk.Frame(card_container, bg="white", highlightbackground="#cccccc", highlightthickness=1, width=300, height=350)
        reimb_card.pack(side="left", padx=20)
        reimb_card.pack_propagate(False)

        tk.Label(reimb_card, text="📝", font=("Arial", 50), bg="white", fg=THEME_COLOR).pack(pady=(40, 10))
        tk.Label(reimb_card, text="General Vouchers", font=("Arial", 16, "bold"), bg="white", fg=TEXT_COLOR).pack(pady=5)
        tk.Label(reimb_card, text="Process employee expenses,\nclaims, and travel reimbursements.", 
                    font=("Arial", 10), fg="#666666", bg="white").pack(pady=5)

        tk.Button(reimb_card, text="Start Processing", command=self.start_reimbursement,
                     bg=THEME_COLOR, fg="white", font=("Arial", 11, "bold"), relief="flat", padx=20, pady=5).pack(side="bottom", pady=40)

        voucher_card = tk.Frame(card_container, bg="white", highlightbackground="#cccccc", highlightthickness=1, width=300, height=350)
        voucher_card.pack(side="left", padx=20)
        voucher_card.pack_propagate(False)

        tk.Label(voucher_card, text="📦", font=("Arial", 50), bg="white", fg="#D84315").pack(pady=(40, 10))
        tk.Label(voucher_card, text="Job Vouchers", font=("Arial", 16, "bold"), bg="white", fg=TEXT_COLOR).pack(pady=5)
        tk.Label(voucher_card, text="Process C&F charges and vendor vouchers.", 
                    font=("Arial", 10), fg="#666666", bg="white").pack(pady=5)

        tk.Button(voucher_card, text="Start Processing", command=self.start_voucher,
                     bg="#D84315", fg="white", font=("Arial", 11, "bold"), relief="flat", padx=20, pady=5).pack(side="bottom", pady=40)

    def start_reimbursement(self):
        self.current_mode = "REIMBURSEMENT"
        self.setup_processing_ui()

    def start_voucher(self):
        self.current_mode = "VOUCHER"
        self.setup_processing_ui()

    def back_to_home(self):
        self.current_mode = None
        self.reimbursement_data = None
        self.expense_details = None
        self.setup_selection_ui()

    def setup_processing_ui(self):
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        self.generate_btn.pack(side="right")

        self.sidebar = ttk.Frame(self.content_frame, width=250, padding=10, relief="solid", borderwidth=1)
        self.sidebar.pack(side="left", fill="y", padx=(0, 20))

        display_title = "GENERAL VOUCHERS" if self.current_mode == "REIMBURSEMENT" else "JOB VOUCHERS"
        ttk.Label(self.sidebar, text=f"{display_title}\nTOOL", font=("Arial", 14, "bold"), foreground=THEME_COLOR, justify="center").pack(pady=20)

        ttk.Button(self.sidebar, text="← Back to Home", command=self.back_to_home).pack(pady=(0, 20), fill="x")

        self.upload_btn = ttk.Button(self.sidebar, text="Upload Files", command=self.upload_files)
        self.upload_btn.pack(pady=10, fill="x")

        self.crdr_frame = ttk.Frame(self.sidebar)
        self.crdr_frame.pack(pady=10, fill="x")
        ttk.Label(self.crdr_frame, text="Default CR/DR:", font=("Arial", 9, "bold")).pack(anchor="w")
        self.crdr_dropdown = ttk.Combobox(self.crdr_frame, values=["DR", "CR"], textvariable=self.selected_cr_dr, state="readonly")
        self.crdr_dropdown.pack(fill="x", pady=5)

        self.fy_frame = ttk.Frame(self.sidebar)
        self.fy_frame.pack(pady=10, fill="x")
        ttk.Label(self.fy_frame, text="Fiscal Year:", font=("Arial", 9, "bold")).pack(anchor="w")
        self.fy_dropdown = ttk.Combobox(self.fy_frame, values=["25-26", "26-27", "27-28", "28-29", "29-30"], 
                                         textvariable=self.fiscal_year, state="readonly")
        self.fy_dropdown.pack(fill="x", pady=5)
        self.fy_dropdown.bind("<<ComboboxSelected>>", lambda e: [self.save_settings(e), self.refresh_previews(e)])

        self.master_section = ttk.Frame(self.sidebar)
        self.master_section.pack(pady=20, fill="x")

        ttk.Label(self.master_section, text="Master Purpose Code", font=("Arial", 10, "bold")).pack(pady=(0, 5))
        
        all_master_codes = sorted(list(set(self.exp_codes + self.warehouse_codes + self.voucher_codes)))
        self.master_combo = SearchableEntry(self.master_section, values=all_master_codes, command=self.reset_fill_down)
        self.master_combo.set("") 
        self.master_combo.pack(pady=5, fill="x")
        
        ttk.Label(self.master_section, text="Select a code above, then click\nany transaction row to assign it.", 
                 font=("Arial", 9), foreground="gray", wraplength=180).pack(pady=5)

        self.status_label = ttk.Label(self.sidebar, text="Ready", foreground="gray")
        self.status_label.pack(side="bottom", pady=20)

        right_frame = ttk.Frame(self.content_frame)
        right_frame.pack(side="right", fill="both", expand=True)

        ttk.Label(right_frame, text="Processed Data Review", font=("Arial", 12, "bold")).pack(anchor="w", pady=(0, 10))

        self.main_frame_container = ScrollableFrame(right_frame)
        self.main_frame_container.pack(fill="both", expand=True)
        self.main_frame = self.main_frame_container.scrollable_frame

    def format_date(self, date_val):
        if pd.isna(date_val): return "No-Date"
        try:
            if isinstance(date_val, str):
                date_val = pd.to_datetime(date_val)
            return date_val.strftime('%d - %b - %Y')
        except:
            return str(date_val)

    def upload_files(self):
        default_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        files = filedialog.askopenfilenames(
            title="Select file(s)", 
            initialdir=default_dir,
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not files: return

        try:
            all_reimb_dfs = []
            all_details_dfs = []
            
            main_sheet_candidates = ["All_General_Vouchers", "GGN_Warehouse_General_Vouchers"]
            details_sheet_name = "Expense Details"
            
            if self.current_mode == "VOUCHER":
                main_sheet_candidates = ["All_Job_Related_Vouchers"]

            for f in files:
                xl = pd.ExcelFile(f)
                sheet_names = xl.sheet_names
                
                # Search for any of the valid main sheet candidates
                for candidate in main_sheet_candidates:
                    if candidate in sheet_names:
                        all_reimb_dfs.append(xl.parse(candidate))
                        break

                if details_sheet_name in sheet_names:
                    all_details_dfs.append(xl.parse(details_sheet_name))
            
            if not all_reimb_dfs:
                expected_sheets = " or ".join([f"'{s}'" for s in main_sheet_candidates])
                messagebox.showerror("Error", f"No {expected_sheets} sheet found.")
                return

            self.reimbursement_data = pd.concat(all_reimb_dfs).drop_duplicates(subset=['ID'])
            self.expense_details = pd.concat(all_details_dfs).drop_duplicates(subset=['SUBFORM LINK ID'])
            
            self.display_data()
            self.generate_btn.configure(state="normal")
            self.status_label.configure(text=f"Loaded {len(self.reimbursement_data)} IDs", foreground="green")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load: {e}")

    def suggest_code(self, description, codes_list=None):
        if codes_list is None:
            codes_list = self.exp_codes
        if len(codes_list) == 1:
            return codes_list[0]
        desc_upper = str(description).upper()
        
        category_map = {
            "TRAVELLING EXPENSES": ["FLIGHT", "PLANE", "AERO", "AIRLINE", "TICKET", "TRAIN", "RAILWAY", "BUS", "TOUR", "TRAVEL", "LONG DISTANCE", "STAY", "HOTEL", "LODGING", "CFS", "PORT", "JNPT", "ICD", "WAREHOUSE", "SITE", "CUSTOM", "TERMINAL", "CAB", "TAXI", "AUTO", "RICKSHAW", "UBER", "OLA", "RAPIDO", "INDRIVE", "CONVEYANCE", "LOCAL", "METRO"],
            "CONVEYANCE EXPENSES": ["TOLL", "PARKING"],
            "POSTAGE AND COURIER": ["COURIER", "POST", "DHL", "BLUE DART", "STAMP", "SPEEDPOST", "TRACKON"],
            "PETROL AND DIESEL": ["PETROL", "DIESEL", "FUEL", "GAS", "CNG", "HPCL", "BPCL", "IOCL"],
            "WATER CHARGES": ["WATER", "AQUAFINA", "BISLERI", "KINLEY", "JAR"],
            "PRINTING AND STATIONERY": ["STATIONERY", "STATIONARY", "PRINT", "PAPER", "PEN", "MARKER", "ENVELOPE"],
            "XEROX CHARGES": ["XEROX", "PHOTOCOPY", "SCAN", "PRINTOUT"],
            "STAFF WELFARE EXPENSES": ["FOOD", "TEA", "LUNCH", "DINNER", "MEAL", "SNACKS", "SWIGGY", "ZOMATO"],
            "REPAIRS AND MAINTENANCE": ["REPAIR", "MAINTENANCE", "SERVICE", "AC ", "PLUMBING"],
            "OFFICE EXPENSES": ["OFFICE", "CLEANING", "SOAP", "TISSUE", "MOP", "BROOM"]
        }

        if " TO " in desc_upper or " - " in desc_upper:
            for actual_code in codes_list:
                if "TRAVELLING EXPENSES" in actual_code.upper():
                    return actual_code

        for code_target, keywords in category_map.items():
            for kw in keywords:
                if kw in desc_upper:
                    for actual_code in codes_list:
                        if code_target.upper() in actual_code.upper():
                            return actual_code
        for fallback_name in ["OFFICE EXPENSES", "GENERAL EXPENSES", "MISC", "OTHER"]:
            for actual_code in codes_list:
                if fallback_name in actual_code.upper():
                    return actual_code
        return codes_list[0] if codes_list else ""

    def get_row_logic(self, mode, input_branch, job_no_val, lob_selection="IMP"):
        input_branch = str(input_branch).strip().upper()
        output_branch = "HO"
        lob = "GEN"

        if mode == "VOUCHER":
            if "ANDHERI" in input_branch or "JNPT" in input_branch:
                output_branch = "HO"
                lob = "CCL EXP" if lob_selection == "EXP" else "CCL IMP"
            elif "GGN" in input_branch or "HARYANA" in input_branch:
                output_branch = "HARYANA"
                lob = "CCL EXP" if lob_selection == "EXP" else "CCL IMP"
            elif "GUJARAT" in input_branch:
                output_branch = "GUJARAT"
                lob = "CCL EXP" if lob_selection == "EXP" else "CCL IMP"
        elif mode == "REIMBURSEMENT":
            if "GGN WAREHOUSE" in input_branch:
                output_branch = "HARYANA"
                lob = "GEN"
            elif "GUJ WAREHOUSE" in input_branch:
                output_branch = "GUJARAT"
                lob = "GEN"
            elif "GUJARAT" in input_branch: 
                output_branch = "GUJARAT"
                lob = "GEN"
                
        cost_center, formatted_job_no = self.get_lob_details(lob, output_branch, job_no_val, mode)
        return output_branch, lob, cost_center, formatted_job_no

    def get_lob_details(self, lob, branch, job_val, mode=None):
        # Image Logic for Cost Centre
        if "IMP" in lob: cc = "CCL Import"
        elif "EXP" in lob: cc = "CCL Export"
        elif "GEN" in lob:
            if branch == "HO":
                cc = "" if mode == "VOUCHER" else "General"
            else:
                cc = "Warehouse"
        else: cc = "General"
        
        fiscal_year = self.fiscal_year.get()
        
        # Pad job_val to 4 digits if it's numeric for Logisys compatibility
        job_val_str = str(job_val).strip()
        if job_val_str.isdigit():
            job_val_str = job_val_str.zfill(4)
        
        job_fmt = job_val_str
        
        # Branch Codes from Image: 
        # HARYANA -> GGN (GEN mode), HAR (CCL mode)
        # GUJARAT -> GUJ
        # CHENNAI -> MAA
        br_code_ccl = ""
        br_code_gen = ""
        if branch == "HARYANA" or "HAR" in branch or "GGN" in branch:
            br_code_ccl = "HAR"; br_code_gen = "GGN"
        elif "GUJ" in branch:
            br_code_ccl = "GUJ"; br_code_gen = "GUJ"
        elif "MAA" in branch or "CHENNAI" in branch:
            br_code_ccl = "MAA"; br_code_gen = "MAA"
        
        if branch == "HO":
            if lob == "GEN": 
                job_fmt = f"HO/Gen/{job_val_str}/{fiscal_year}"
            else:
                prefix = "ER" if "EXP" in lob else "IR"
                job_fmt = f"{prefix}/{job_val_str}/{fiscal_year}"
        else:
            if lob == "GEN":
                code = br_code_gen or br_code_ccl
                job_fmt = f"WH/{code}/{job_val_str}/{fiscal_year}" if code else f"WH/{job_val_str}/{fiscal_year}"
            else:
                prefix = "ER" if "EXP" in lob else "IR"
                job_fmt = f"{prefix}/{br_code_ccl}/{job_val_str}/{fiscal_year}" if br_code_ccl else f"{prefix}/{job_val_str}/{fiscal_year}"
                
        return cc, job_fmt

    def process_fill_down(self, current_combo, all_combos):
        val = current_combo.get()
        # Track sequential clicks on the same row to push value further down
        if not hasattr(self, '_last_fill_source') or self._last_fill_source != current_combo:
            self._last_fill_source = current_combo
            self._last_fill_offset = 1
        else:
            self._last_fill_offset += 1
            
        try:
            start_idx = all_combos.index(current_combo)
            target_idx = start_idx + self._last_fill_offset
            if target_idx < len(all_combos):
                all_combos[target_idx].set(val)
            else:
                # Reset if we reach the end of the report
                self._last_fill_offset = 0
        except ValueError:
            pass

    def reset_fill_down(self, *args):
        """Resets the sequential tracking for the ↓ button."""
        self._last_fill_source = None
        self._last_fill_offset = 0

    def display_data(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        self.selected_codes = {}
        self.lob_selection = {} 
        self.row_ui_refs = {} 
        self.all_row_combos = []  # Track every single row combo in order for global fill-down

        BRANCH_LOBS = {
            "HO": ["CCL IMP", "CCL EXP", "GEN"],
            "HARYANA": ["GEN", "CCL IMP"],
            "GUJARAT": ["GEN", "CCL IMP"],
            "CHENNAI": ["CCL IMP"]
        }

        def on_row_lob_change(event, sub_id):
            if sub_id not in self.row_ui_refs: return
            lob_val = event.widget.get()
            refs = self.row_ui_refs[sub_id]
            new_cc, new_job_fmt = self.get_lob_details(lob_val, refs['branch'], refs['job_val'], self.current_mode)
            if refs.get('cc_label'): refs['cc_label'].configure(text=f"CC: {new_cc}")
            if refs.get('badge_label'): refs['badge_label'].configure(text=f"[{new_job_fmt}]")

        for _, row in self.reimbursement_data.iterrows():
            # Robustly get the transaction ID/Voucher number
            trans_id = row.get('Voucher Number', row.get('Transaction ID', ''))
            formatted_date = self.format_date(row['Payment Date'])
            parent_id = row['ID']
            emp_name = row['Employee Name']
            input_branch = row.get('Branch', '')
            
            group_frame = ttk.Frame(self.main_frame, style="Card.TFrame", padding=5, relief="solid", borderwidth=1)
            group_frame.pack(fill="x", pady=10, padx=10)
            
            header = ttk.Frame(group_frame)
            header.pack(fill="x", padx=5, pady=5)
            
            header_text = f"Voucher Number: {trans_id}  |  {emp_name}"
            if self.current_mode != "VOUCHER":
                header_text += f"  |  {formatted_date}"
            
            ttk.Label(header, text=header_text, font=("Arial", 11, "bold"), foreground=THEME_COLOR).pack(side="left")
           
            trans_combos = [] 
            
            # Determine if this is a Warehouse Transaction (Job-related)
            is_trans_warehouse = False
            
            # 1. Check Parent row (often contains the JOB NO)
            parent_norm = {str(k).strip().upper(): v for k, v in row.items()}
            for col_variant in ['JOB NO', 'JOBNO', 'JOB ID', 'JOB', 'JOB NUMBER', 'JOBS']:
                val = parent_norm.get(col_variant)
                if val is not None and str(val).strip() != "" and str(val).lower() != "nan":
                    is_trans_warehouse = True; break
            
            # 2. Check Subform details
            if not is_trans_warehouse:
                details_check = self.expense_details[self.expense_details['PARENT ID'] == parent_id]
                for _, d_row in details_check.iterrows():
                    norm_row = {str(k).strip().upper(): v for k, v in d_row.items()}
                    for col_variant in ['JOB NO', 'JOBNO', 'JOB ID', 'JOB', 'JOB NUMBER', 'JOBS']:
                            val = norm_row.get(col_variant)
                            if val is not None and str(val).strip() != "" and str(val).lower() != "nan":
                                is_trans_warehouse = True; break
                    if is_trans_warehouse: break

            if self.current_mode == "VOUCHER": 
                current_codes = self.voucher_codes
            else: 
                current_codes = self.warehouse_codes if is_trans_warehouse else self.exp_codes 

            bulk_apply_frame = ttk.Frame(header)
            bulk_apply_frame.pack(side="right")
            ttk.Label(bulk_apply_frame, text="Bulk Apply:", font=("Arial", 9, "italic"), foreground="gray").pack(side="left", padx=5)
            bulk_apply = SearchableEntry(bulk_apply_frame, values=current_codes, 
                                        command=lambda val, tc=trans_combos: [self.reset_fill_down(), [e.set(val) for e in tc]], width=250)
            bulk_apply.pack(side="left")

            details = self.expense_details[self.expense_details['PARENT ID'] == parent_id]
            for _, d_row in details.iterrows():
                sub_id = d_row['SUBFORM LINK ID']
                desc = d_row.get('Item Description' if self.current_mode == "VOUCHER" else 'Expense Description', '')
                amt = d_row.get('Expense Amount', 0)

                row_frame = ttk.Frame(group_frame, style="Row.TFrame")
                row_frame.pack(fill="x", pady=5, padx=10)
                
                row_frame.grid_columnconfigure(3, weight=5) 
                
                def find_job(data_row):
                    norm_row = {str(k).strip().upper(): v for k, v in data_row.items()}
                    # Check all variants for both modes to be safe
                    for col in ['JOB NO', 'JOBNO', 'JOB ID', 'JOB', 'JOB NUMBER', 'JOBS']:
                        val = norm_row.get(col)
                        if val is not None and str(val).strip() != "" and str(val).lower() != "nan":
                            s = str(val).strip()
                            return s[:-2] if s.endswith(".0") else s
                    return ""

                job_val = find_job(d_row) or find_job(row)
                out_branch, init_lob, init_cc, init_fmt_job = self.get_row_logic(self.current_mode, input_branch, job_val, "IMP")
                allowed_lobs = BRANCH_LOBS.get(out_branch, ["GEN"])
                if init_lob not in allowed_lobs and allowed_lobs:
                     init_lob = allowed_lobs[0]
                
                # Get CC for processing
                init_cc, init_fmt_job = self.get_lob_details(init_lob, out_branch, job_val, self.current_mode)

                current_lob_var = tk.StringVar(value=init_lob)
                self.lob_selection[sub_id] = current_lob_var
                
                if job_val:
                    badge_lbl = ttk.Label(row_frame, text=f"[{init_fmt_job}]", foreground=THEME_COLOR, font=("Arial", 8, "bold"), width=22)
                    badge_lbl.grid(row=0, column=0, padx=5, sticky="w")
                    
                    lob_menu = ttk.Combobox(row_frame, values=allowed_lobs, textvariable=current_lob_var, width=10, state="readonly")
                    lob_menu.bind("<<ComboboxSelected>>", lambda e, s=sub_id: on_row_lob_change(e, s))
                    lob_menu.grid(row=0, column=1, padx=4, sticky="w")
                    
                    cc_lbl = ttk.Label(row_frame, text=f"CC: {init_cc}", foreground="gray", width=15)
                    cc_lbl.grid(row=0, column=2, padx=4, sticky="w")
                    
                    self.row_ui_refs[sub_id] = {'branch': out_branch, 'job_val': job_val, 'badge_label': badge_lbl, 'cc_label': cc_lbl}
                else:
                    badge_lbl = ttk.Label(row_frame, text="[NON-JOB]", foreground="#546E7A", font=("Arial", 8, "bold"), width=22)
                    badge_lbl.grid(row=0, column=0, padx=5, sticky="w")
                    
                    cc_text = f"CC: {init_cc}" if self.current_mode == "VOUCHER" else ""
                    cc_lbl = ttk.Label(row_frame, text=cc_text, foreground="gray", width=15)
                    cc_lbl.grid(row=0, column=2, padx=4, sticky="w")
                    self.row_ui_refs[sub_id] = {'branch': out_branch, 'job_val': job_val, 'badge_label': badge_lbl, 'cc_label': cc_lbl}

                desc_lbl = ttk.Label(row_frame, text=desc, font=("Arial", 9), wraplength=450, justify="left")
                desc_lbl.grid(row=0, column=3, padx=10, sticky="ew")
                if len(str(desc)) > 100: ToolTip(desc_lbl, str(desc))

                ttk.Label(row_frame, text=f"Amt: {amt}", width=12, anchor="e").grid(row=0, column=4, padx=5, sticky="e")

                combo = SearchableEntry(row_frame, values=current_codes, width=300, command=self.reset_fill_down)
                copy_btn = ttk.Button(row_frame, text="↓", width=3, command=lambda c=combo, l=self.all_row_combos: self.process_fill_down(c, l))
                copy_btn.grid(row=0, column=5, padx=2)
                combo.grid(row=0, column=6, padx=10, pady=5, sticky="e")
                
                self.selected_codes[sub_id] = {'combo': combo, 'row_data': d_row}
                trans_combos.append(combo) 
                self.all_row_combos.append(combo) 
                
                suggestion = self.suggest_code(desc, current_codes)
                if suggestion: combo.set(suggestion)
                elif len(current_codes) == 1: combo.set(current_codes[0])

                def apply_master(event, c=combo):
                    val = self.master_combo.get().strip()
                    if val: c.set(val)

                row_frame.bind("<Button-1>", apply_master)
                for child in row_frame.winfo_children(): 
                    if child == copy_btn: continue # Don't apply master on copy button click
                    child.bind("<Button-1>", apply_master)

            # After all rows for this transaction are added, pre-fill for Voucher
            if self.current_mode == "VOUCHER":
                bulk_apply.set(self.voucher_codes[0], trigger_command=True)

    def generate_output(self):
        # Let the user choose the output directory
        output_dir = filedialog.askdirectory(title="Select Folder to Save Output Files")
        if not output_dir:
            return  # User cancelled folder selection
        
        # Scan for existing Voucher Numbers in the persistent history log
        history = self.load_history()
        
        duplicate_transactions = []
        
        # Check each transaction for history matches
        for _, row in self.reimbursement_data.iterrows():
            trans_id_str = str(row.get('Voucher Number', row.get('Transaction ID', '')))
            
            # If ID exists in history for the current mode, it's a duplicate
            if trans_id_str in history:
                # We also check the mode to be extra safe, though IDs are usually unique
                if history[trans_id_str].get("mode") == self.current_mode:
                    if trans_id_str not in duplicate_transactions:
                        duplicate_transactions.append(trans_id_str)
        
        # Show warning if duplicates found in history
        skip_duplicates = False
        if duplicate_transactions:
            duplicate_list = "\n".join([f"  • {tid}" for tid in duplicate_transactions])
            mode_name = "GENERAL VOUCHER" if self.current_mode == "REIMBURSEMENT" else "JOB VOUCHER"
            warning_message = (f"⚠️ WARNING: {mode_name} output files already exist for the following Voucher Numbers:\n\n"
                               f"{duplicate_list}\n\n"
                               "What would you like to do?\n"
                               "• Click [Yes] to OVERWRITE all existing files.\n"
                               "• Click [No] to SKIP duplicates and only generate new files.\n"
                               "• Click [Cancel] to ABORT.")
            
            response = messagebox.askyesnocancel("Duplicate Output Files Detected", warning_message, icon='warning')
            if response is None:
                return  # User clicked Cancel or closed the dialog
            if response is False:
                skip_duplicates = True
            # If response is True, skip_duplicates remains False (proceed with overwrite)
        
        missing_count = 0
        for sub_id, entry_dict in self.selected_codes.items():
            if not entry_dict['combo'].get().strip(): missing_count += 1
        if missing_count > 0:
            messagebox.showwarning("Mandatory Fields", f"Validation Failed: {missing_count} Expense Codes are missing.")
            return
        
        try:
            cr_dr = self.selected_cr_dr.get()
            logisys_cols = ['Charge/GL', 'Charge/GL Name', 'Charge/GL Amount', 'Branch', 'LOB', 'JobType', 'JobNo', 'Amount', 'CR/DR', 'Cost Center', 'Narration', 'Start Date', 'End Date', 'AWB/HAWB/MBL/HBL No', 'WHTaxOrgName', 'WHTaxCode', 'WHTaxTaxableAmt', 'WHTaxRate', 'CC Code']
            warehouse_cols = ['LedgerName', 'AmountHC', 'AmtType', 'Branch', 'CostCenter', 'Remarks', 'ChargeName', 'LOB', 'JobType', 'JobNo', 'AmountFC', 'PeriodStartDate', 'PeriodEndDate', 'WHTaxOrgName', 'WHTaxCode', 'WHTaxTaxableAmt', 'WHTaxRate']

            generated_count = 0
            for _, row in self.reimbursement_data.iterrows():
                trans_id = row.get('Voucher Number', row.get('Transaction ID', ''))
                fmt_date, parent_id = self.format_date(row['Payment Date']), row['ID']
                
                # Logic to skip duplicates if requested
                if skip_duplicates and str(trans_id) in duplicate_transactions:
                    continue
                
                input_branch = row.get('Branch', '')
                
                details_for_check = self.expense_details[self.expense_details['PARENT ID'] == parent_id]
                is_trans_warehouse = (self.current_mode == "VOUCHER")
                if not is_trans_warehouse:
                    # Check parent row
                    parent_norm = {str(k).strip().upper(): v for k, v in row.items()}
                    for col_variant in ['JOB NO', 'JOBNO', 'JOB ID', 'JOB', 'JOB NUMBER', 'JOBS']:
                        val = parent_norm.get(col_variant)
                        if val is not None and str(val).strip() != "" and str(val).lower() != "nan":
                            is_trans_warehouse = True; break
                    
                    # Check details if parent didn't have it
                    if not is_trans_warehouse:
                        for _, d_row in details_for_check.iterrows():
                            norm_row = {str(k).strip().upper(): v for k, v in d_row.items()}
                            for col_variant in ['JOB NO', 'JOBNO', 'JOB ID', 'JOB', 'JOB NUMBER', 'JOBS']:
                                    val = norm_row.get(col_variant)
                                    if val is not None and str(val).strip() != "" and str(val).lower() != "nan":
                                        is_trans_warehouse = True; break
                            if is_trans_warehouse: break
                
                output_rows = []
                for _, d_row in details_for_check.iterrows():
                    sub_id = d_row['SUBFORM LINK ID']
                    selected_val = self.selected_codes[sub_id]['combo'].get()
                    amt = d_row.get('Expense Amount', 0)
                    desc = d_row.get('Item Description' if self.current_mode == "VOUCHER" else 'Expense Description', '')
                    
                    def get_job_val(row_data, parent_data):
                        # Combine dictionaries and normalize keys to uppercase
                        combined = {**{str(k).strip().upper(): v for k, v in parent_data.items()},
                                   **{str(k).strip().upper(): v for k, v in row_data.items()}}
                        for c in ['JOB NO', 'JOBNO', 'JOB ID', 'JOB', 'JOB NUMBER', 'JOBS']:
                            val = combined.get(c)
                            if val is not None and str(val).strip() != "" and str(val).lower() != "nan":
                                s = str(val).strip()
                                return s[:-2] if s.endswith(".0") else s
                        return ""

                    job_val = get_job_val(d_row, row)

                    # Always calculate branch and lob logic, even if job_val is empty
                    lob = self.lob_selection[sub_id].get() if sub_id in self.lob_selection else "GEN"
                    out_branch, _, _, _ = self.get_row_logic(self.current_mode, input_branch, job_val, lob)
                    out_cost_center, job_no = self.get_lob_details(lob, out_branch, job_val, self.current_mode)
                    
                    if is_trans_warehouse:
                        csv_row = {c: None for c in warehouse_cols}
                        csv_row.update({
                            'LedgerName': selected_val, 
                            'AmountHC': amt, 
                            'AmtType': "DEBIT" if self.current_mode == "VOUCHER" else cr_dr, 
                            'Branch': out_branch, 
                            'CostCenter': out_cost_center, 
                            'Remarks': desc, 
                            'ChargeName': "C&F CHARGES" if self.current_mode == "VOUCHER" else "", 
                            'LOB': lob, 
                            'JobType': '',
                            'JobNo': job_no, 
                            'AmountFC': amt
                        })
                        output_rows.append([csv_row[c] for c in warehouse_cols])
                    else:
                        out_row = {c: None for c in logisys_cols}
                        out_row.update({
                            'Charge/GL': selected_val, 
                            'Charge/GL Name': selected_val, 
                            'Charge/GL Amount': amt, 
                            'Amount': amt, 
                            'Branch': out_branch, 
                            'CR/DR': cr_dr, 
                            'Narration': desc, 
                            'LOB': '',
                            'JobType': '',
                            'JobNo': '',
                            'Cost Center': '',
                            'CC Code': ''
                        })
                        output_rows.append([out_row[c] for c in logisys_cols])

                if self.current_mode == "VOUCHER":
                    filename = f"{row['Employee Name']} JOB VOUCHER NO. {trans_id}.csv"
                    pd.DataFrame(output_rows, columns=warehouse_cols).to_csv(os.path.join(output_dir, filename), index=False)
                elif is_trans_warehouse:
                    filename = f"{row['Employee Name']} GENERAL VOUCHER NO. {trans_id}.csv"
                    pd.DataFrame(output_rows, columns=warehouse_cols).to_csv(os.path.join(output_dir, filename), index=False)
                else:
                    filename = f"{row['Employee Name']} GENERAL VOUCHER NO. {trans_id}.xlsx"
                    wb = Workbook()
                    ws = wb.active
                    ws.append(logisys_cols)
                    for r_data in output_rows: ws.append(r_data)
                    for cell in ws[1]: cell.font = Font(bold=True)
                    wb.save(os.path.join(output_dir, filename))
                
                # Mark as processed in the history log
                self.save_to_history(trans_id, self.current_mode)
                generated_count += 1
            
            if generated_count == 0 and skip_duplicates:
                messagebox.showinfo("Process Complete", "No new files were generated. All items were skipped as they already exist.")
            else:
                messagebox.showinfo("Success", f"Generated {generated_count} outputs in:\n{output_dir}")
                
            os.startfile(output_dir)
        except Exception as e:
            messagebox.showerror("Error", f"Failed generating output: {e}")

if __name__ == "__main__":
    ReimbursementApp().mainloop()

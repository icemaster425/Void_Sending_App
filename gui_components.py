import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font
import os
from datetime import datetime

class MainWindow:
    def __init__(self, root, app_controller):
        self.root = root
        self.app = app_controller
        
        self.current_user = "Pending..."
        self.root.title(f"V.O.I.D. - Verified On-boarding Institutional Dispatcher [User: {self.current_user}]")
        
        self.setup_styles()
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.main_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.main_tab, text="Monitoring")
        self.setup_monitoring_tab(self.main_tab)
        
        self.institutions_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.institutions_tab, text="Institutions")
        self.setup_institutions_tab(self.institutions_tab)

        self.settings_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_tab, text="Settings")
        self.setup_settings_tab(self.settings_tab)
        
        self.batch_panels = {}
        
        try:
            self.load_institutions()
            self.load_processed_batches()
        except Exception:
            pass 

        self.root.after(200, self.show_login_popup)

    def show_login_popup(self):
        staff_fallback = "Dianne, Dony, Natasha, Nes, Ricci, Lynda, Maria, Michaela"
        try:
            staff_str = self.app.master_config.get('SHARED_SETTINGS', 'staff_list', fallback=staff_fallback)
        except AttributeError:
            staff_str = staff_fallback
            
        staff_list = [s.strip() for s in staff_str.split(',')]
        
        last_user = self.app.local_config.get('PREFS', 'last_user', fallback="") if self.app.local_config.has_section('PREFS') else ""
        default_user = last_user if last_user in staff_list else staff_list[0]
        
        win = tk.Toplevel(self.root)
        win.title("V.O.I.D. Login")
        win.geometry("300x150")
        win.resizable(False, False)
        
        win.transient(self.root)          
        win.attributes('-topmost', True)  
        
        win.update_idletasks()
        x = (win.winfo_screenwidth() // 2) - (win.winfo_width() // 2)
        y = (win.winfo_screenheight() // 2) - (win.winfo_height() // 2)
        win.geometry(f'+{x}+{y}')

        tk.Label(win, text="Who is dispatching today?", font=("Helvetica", 11, "bold"), bg=win.cget("bg")).pack(pady=15)
        
        user_var = tk.StringVar(value=default_user)
        combo = ttk.Combobox(win, textvariable=user_var, values=staff_list, state="readonly", font=("Helvetica", 10))
        combo.pack(pady=5, padx=30, fill=tk.X)
        
        def on_confirm():
            selected_name = user_var.get()
            self.current_user = selected_name
            self.root.title(f"V.O.I.D. - Verified On-boarding Institutional Dispatcher [User: {self.current_user}]")
            
            if not self.app.local_config.has_section('PREFS'):
                self.app.local_config.add_section('PREFS')
            self.app.local_config.set('PREFS', 'last_user', selected_name)
            try:
                with open(self.app.local_config_path, 'w') as f:
                    self.app.local_config.write(f)
            except Exception:
                pass 
            win.destroy()
            
        ttk.Button(win, text="Start System", command=on_confirm).pack(pady=15)
        win.protocol("WM_DELETE_WINDOW", lambda: self.root.destroy()) 
        win.grab_set()
        win.focus_force()

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("alt")
        purple, gold = "#4B0082", "#FFD700"
        
        self.style.configure("TNotebook", background=purple)
        self.style.configure("TNotebook.Tab", padding=[10, 5], background=gold, foreground="black")
        self.style.map("TNotebook.Tab", background=[("selected", "#FFA500")])
        self.style.configure("TFrame", background="#F0F0F0")
        self.style.configure("TLabelFrame", background="#F0F0F0", foreground=purple, font=("Helvetica", 12, "bold"))
        self.style.configure("Treeview", rowheight=25)
        self.style.map("Treeview", background=[("selected", "#3498db")], foreground=[("selected", "white")])
        self.style.configure("Glossy.TButton", font=("Helvetica", 10, "bold"), borderwidth=2, padding=[10, 5])
        self.style.configure("Add.Glossy.TButton", background="#28a745", foreground="white")
        self.style.configure("Edit.Glossy.TButton", background="#007bff", foreground="white")
        self.style.configure("Delete.Glossy.TButton", background="#dc3545", foreground="white")
        self.style.configure("Multi.Glossy.TButton", background="#fd7e14", foreground="white")
        
        # --- EXPLICIT UNIFIED NORMAL STYLES ---
        normal_bg = "#E6E6E6" 
        self.style.configure("Normal.TLabelframe", background=normal_bg)
        self.style.configure("Normal.TLabelframe.Label", background=normal_bg, foreground=purple, font=("Helvetica", 12, "bold"))
        self.style.configure("Normal.TFrame", background=normal_bg)
        self.style.configure("Normal.TCheckbutton", background=normal_bg)
        self.style.configure("NormalText.TLabel", background=normal_bg, foreground="black", font=("Helvetica", 10, "bold"))

        # --- EXPLICIT UNIFIED WARNING STYLES ---
        warn_bg = "#ffcccc"
        self.style.configure("Warning.TLabelframe", background=warn_bg)
        self.style.configure("Warning.TLabelframe.Label", background=warn_bg, foreground="red", font=("Helvetica", 12, "bold"))
        self.style.configure("Warning.TFrame", background=warn_bg)
        self.style.configure("Warning.TCheckbutton", background=warn_bg)
        self.style.configure("WarningText.TLabel", background=warn_bg, foreground="red", font=("Helvetica", 10, "bold"))

    def setup_monitoring_tab(self, parent_frame):
        self.folder_frame = ttk.LabelFrame(parent_frame, text="Folder Monitoring", padding="10")
        self.folder_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.batches_main_frame = ttk.Frame(parent_frame)
        self.batches_main_frame.pack(fill=tk.BOTH, expand=True)

        self.detected_batches_frame = ttk.LabelFrame(self.batches_main_frame, text="Detected Batches", padding="10")
        self.detected_batches_frame.pack(fill=tk.BOTH, expand=True, padx=(0, 10), side=tk.LEFT)
        
        self.processed_batches_frame = ttk.LabelFrame(self.batches_main_frame, text="Processed Batches (Shared)", padding="10")
        self.processed_batches_frame.pack(fill=tk.BOTH, expand=True, side=tk.RIGHT)

        self.filter_frame = ttk.Frame(self.processed_batches_frame)
        self.filter_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(self.filter_frame, text="Find (Batch/Inst/Date):").pack(side=tk.LEFT, padx=2)
        self.search_entry = ttk.Entry(self.filter_frame)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.search_entry.bind("<KeyRelease>", lambda e: self.filter_history())
        ttk.Button(self.filter_frame, text="Today", style="Edit.Glossy.TButton", command=self.show_today_only).pack(side=tk.LEFT, padx=2)
        ttk.Button(self.filter_frame, text="All", style="Glossy.TButton", command=self.load_processed_batches).pack(side=tk.LEFT, padx=2)

        self.folder_path_entry = ttk.Entry(self.folder_frame, width=40)
        self.folder_path_entry.grid(row=0, column=0, padx=5, pady=5)
        ttk.Button(self.folder_frame, text="Browse", style="Glossy.TButton", command=self.browse_folder).grid(row=0, column=1, padx=5, pady=5)
        self.start_button = ttk.Button(self.folder_frame, text="Start", style="Glossy.TButton", command=self.start_monitoring)
        self.start_button.grid(row=0, column=2, padx=5, pady=5)
        self.stop_button = ttk.Button(self.folder_frame, text="Stop", style="Glossy.TButton", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=3, padx=5, pady=5)
        ttk.Button(self.folder_frame, text="Clear Folder", style="Delete.Glossy.TButton", command=self.app.delete_all_files).grid(row=0, column=4, padx=5, pady=5)

        self.batches_canvas = tk.Canvas(self.detected_batches_frame)
        self.batches_scrollbar = ttk.Scrollbar(self.detected_batches_frame, orient="vertical", command=self.batches_canvas.yview)
        self.batches_inner_frame = ttk.Frame(self.batches_canvas)
        self.batches_inner_frame.bind("<Configure>", lambda e: self.batches_canvas.configure(scrollregion=self.batches_canvas.bbox("all")))
        self.batches_canvas.create_window((0, 0), window=self.batches_inner_frame, anchor="nw")
        self.batches_canvas.configure(yscrollcommand=self.batches_scrollbar.set)
        self.batches_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.batches_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.processed_tree = ttk.Treeview(self.processed_batches_frame, columns=("inst", "batch", "date", "time", "files", "user", "recs"), show="headings")
        headings = ["Inst", "Batch #", "Date", "Time", "Files", "Dispatcher", "Recs"]
        for col, head in zip(self.processed_tree["columns"], headings):
            self.processed_tree.heading(col, text=head, command=lambda c=col: self.sort_column(c, False))
        self.processed_tree.pack(fill=tk.BOTH, expand=True)
        
        self.log_frame = ttk.LabelFrame(parent_frame, text="Activity Log", padding="10")
        self.log_frame.pack(fill=tk.X, pady=(10, 0))
        self.log_text = tk.Text(self.log_frame, state="disabled", wrap=tk.WORD, height=5)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def autofit_columns(self):
        padding = 20
        measure_font = font.Font() 
        for col in self.processed_tree["columns"]:
            max_w = measure_font.measure(self.processed_tree.heading(col)['text']) + padding
            for item in self.processed_tree.get_children():
                cell_val = str(self.processed_tree.set(item, col))
                cell_w = measure_font.measure(cell_val) + padding
                if cell_w > max_w:
                    max_w = cell_w
            self.processed_tree.column(col, width=max_w)

    def show_today_only(self):
        if not self.app.db_manager: return
        today = datetime.now().strftime('%d/%m/%Y')
        for i in self.processed_tree.get_children(): self.processed_tree.delete(i)
        for b in self.app.db_manager.get_sent_emails_by_date(today):
            self.processed_tree.insert("", "end", values=b)
        self.autofit_columns()

    def filter_history(self):
        if not self.app.db_manager: return
        term = self.search_entry.get()
        for i in self.processed_tree.get_children(): self.processed_tree.delete(i)
        for b in self.app.db_manager.search_sent_emails(term):
            self.processed_tree.insert("", "end", values=b)
        self.autofit_columns()

    def sort_column(self, col, reverse):
        l = [(self.processed_tree.set(k, col), k) for k in self.processed_tree.get_children("")]
        l.sort(reverse=reverse)
        for index, (val, k) in enumerate(l):
            self.processed_tree.move(k, "", index)
        self.processed_tree.heading(col, command=lambda: self.sort_column(col, not reverse))

    def setup_institutions_tab(self, parent_frame):
        tree_frame = ttk.LabelFrame(parent_frame, text="Institution Database", padding="10")
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 5))
        
        self.institutions_tree = ttk.Treeview(tree_frame, columns=("c_code", "i_code", "email", "key", "msg"), show="headings", selectmode="extended")
        headings = ["Country Code", "Inst Code", "Email", "Password", "Message"]
        for col, text in zip(self.institutions_tree["columns"], headings):
            self.institutions_tree.heading(col, text=text)
        self.institutions_tree.pack(fill=tk.BOTH, expand=True)
        
        self.institutions_tree.bind("<<TreeviewSelect>>", self.on_institution_select)
        
        btn_frame = ttk.Frame(tree_frame)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="Add New", style="Add.Glossy.TButton", command=self.add_institution).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Popup Edit", style="Edit.Glossy.TButton", command=self.edit_institution).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Multi-Edit", style="Multi.Glossy.TButton", command=self.multi_edit_institutions).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Delete", style="Delete.Glossy.TButton", command=self.delete_institution).pack(side=tk.LEFT, padx=5)

        self.quick_edit_frame = ttk.LabelFrame(parent_frame, text="Quick Edit (Select 1 Row)", padding="10")
        self.quick_edit_frame.pack(fill=tk.X, padx=10, pady=(5, 10))
        
        self.qe_vars = {
            'c_code': tk.StringVar(),
            'i_code': tk.StringVar(),
            'email': tk.StringVar(),
            'key': tk.StringVar()
        }
        
        ttk.Label(self.quick_edit_frame, text="Country:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.qe_c_code = ttk.Entry(self.quick_edit_frame, textvariable=self.qe_vars['c_code'], width=15)
        self.qe_c_code.grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(self.quick_edit_frame, text="Inst Code:").grid(row=0, column=2, padx=5, pady=2, sticky=tk.W)
        self.qe_i_code = ttk.Entry(self.quick_edit_frame, textvariable=self.qe_vars['i_code'], width=15, state="readonly")
        self.qe_i_code.grid(row=0, column=3, padx=5, pady=2)
        
        ttk.Label(self.quick_edit_frame, text="Email:").grid(row=0, column=4, padx=5, pady=2, sticky=tk.W)
        self.qe_email = ttk.Entry(self.quick_edit_frame, textvariable=self.qe_vars['email'], width=30)
        self.qe_email.grid(row=0, column=5, padx=5, pady=2)
        
        ttk.Label(self.quick_edit_frame, text="Password:").grid(row=0, column=6, padx=5, pady=2, sticky=tk.W)
        self.qe_key = ttk.Entry(self.quick_edit_frame, textvariable=self.qe_vars['key'], width=20)
        self.qe_key.grid(row=0, column=7, padx=5, pady=2)
        
        ttk.Label(self.quick_edit_frame, text="Message:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.NW)
        self.qe_msg = tk.Text(self.quick_edit_frame, width=80, height=3, font=("Segoe UI", 9))
        self.qe_msg.grid(row=1, column=1, columnspan=5, padx=5, pady=5)
        
        self.btn_quick_save = ttk.Button(self.quick_edit_frame, text="Save Quick Edit", style="Edit.Glossy.TButton", command=self.save_quick_edit, state=tk.DISABLED)
        self.btn_quick_save.grid(row=1, column=6, columnspan=2, padx=5, pady=5)

    def on_institution_select(self, event):
        selected_items = self.institutions_tree.selection()
        
        if len(selected_items) == 1:
            self.btn_quick_save.config(state=tk.NORMAL)
            code = self.institutions_tree.item(selected_items[0], 'values')[1]
            data = self.app.db_manager.get_institution_by_code(code)
            
            self.qe_vars['c_code'].set(data.get('county_code', ''))
            
            self.qe_i_code.config(state=tk.NORMAL)
            self.qe_vars['i_code'].set(data.get('institution_code', ''))
            self.qe_i_code.config(state="readonly") 
            
            self.qe_vars['email'].set(data.get('email', ''))
            self.qe_vars['key'].set(data.get('encryption_key', ''))
            
            self.qe_msg.delete("1.0", tk.END)
            self.qe_msg.insert("1.0", data.get('message', ''))
        else:
            self.btn_quick_save.config(state=tk.DISABLED)
            for var in self.qe_vars.values(): var.set('')
            self.qe_msg.delete("1.0", tk.END)

    def save_quick_edit(self):
        code = self.qe_vars['i_code'].get()
        if not code: return
        
        up_data = {
            'county_code': self.qe_vars['c_code'].get(),
            'institution_code': code,
            'email': self.qe_vars['email'].get(),
            'encryption_key': self.qe_vars['key'].get(),
            'message': self.qe_msg.get("1.0", tk.END).strip()
        }
        self.app.db_manager.update_institution(code, up_data)
        self.load_institutions()
        messagebox.showinfo("Success", f"Updated {code} successfully.")
        
        for item in self.institutions_tree.get_children():
            if self.institutions_tree.item(item, 'values')[1] == code:
                self.institutions_tree.selection_set(item)
                break

    def add_institution(self):
        win = tk.Toplevel(self.root)
        win.title("Add New Institution")
        fields = ["Country Code:", "Institution Code:", "Email:", "Encryption Key:", "Message:"]
        ents = {}
        for i, f in enumerate(fields):
            ttk.Label(win, text=f).grid(row=i, column=0, padx=10, pady=5, sticky=tk.NW)
            if f == "Message:":
                ent = tk.Text(win, width=40, height=6, font=("Segoe UI", 9))
                ent.grid(row=i, column=1, padx=10, pady=5)
            else:
                ent = ttk.Entry(win, width=50)
                ent.grid(row=i, column=1, padx=10, pady=5)
            ents[f] = ent
        def save():
            msg = ents["Message:"].get("1.0", tk.END).strip()
            self.app.db_manager.add_institution(ents["Country Code:"].get(), ents["Institution Code:"].get(), ents["Email:"].get(), ents["Encryption Key:"].get(), msg)
            self.load_institutions()
            win.destroy()
        ttk.Button(win, text="Save", style="Add.Glossy.TButton", command=save).grid(row=len(fields), column=0, columnspan=2, pady=10)

    def edit_institution(self):
        sel = self.institutions_tree.selection()
        if not sel: return
        code = self.institutions_tree.item(sel[0], 'values')[1]
        data = self.app.db_manager.get_institution_by_code(code)
        win = tk.Toplevel(self.root)
        win.title(f"Edit: {code}")
        fields = ["Country Code:", "Institution Code:", "Email:", "Encryption Key:", "Message:"]
        ents = {}
        for i, f in enumerate(fields):
            ttk.Label(win, text=f).grid(row=i, column=0, padx=10, pady=5, sticky=tk.NW)
            if f == "Message:":
                ent = tk.Text(win, width=40, height=6, font=("Segoe UI", 9))
                ent.insert("1.0", data.get('message', ''))
                ent.grid(row=i, column=1, padx=10, pady=5)
            else:
                ent = ttk.Entry(win, width=50)
                key_map = {"Country Code:": 'county_code', "Institution Code:": 'institution_code', "Email:": 'email', "Encryption Key:": 'encryption_key'}
                ent.insert(0, data.get(key_map[f], ''))
                ent.grid(row=i, column=1, padx=10, pady=5)
            ents[f] = ent
        def save():
            up_data = {'county_code': ents["Country Code:"].get(), 'institution_code': ents["Institution Code:"].get(), 'email': ents["Email:"].get(), 'encryption_key': ents["Encryption Key:"].get(), 'message': ents["Message:"].get("1.0", tk.END).strip()}
            self.app.db_manager.update_institution(code, up_data)
            self.load_institutions()
            win.destroy()
        ttk.Button(win, text="Save Changes", style="Edit.Glossy.TButton", command=save).grid(row=len(fields), column=0, columnspan=2, pady=10)

    def multi_edit_institutions(self):
        selected_items = self.institutions_tree.selection()
        if len(selected_items) < 2: return
        win = tk.Toplevel(self.root)
        win.title("Multi-Edit Records")
        ttk.Label(win, text="Update selected rows:").grid(row=0, column=0, columnspan=2, pady=10)
        fields = ["Email Domain:", "Encryption Key:", "Message:"]
        ents = {}
        for i, f in enumerate(fields, 1):
            ttk.Label(win, text=f).grid(row=i, column=0, padx=10, pady=5, sticky=tk.NW)
            if f == "Message:":
                ent = tk.Text(win, width=40, height=4)
                ent.grid(row=i, column=1, padx=10, pady=5)
            else:
                ent = ttk.Entry(win, width=40)
                ent.grid(row=i, column=1, padx=10, pady=5)
            ents[f] = ent
        def apply_batch():
            for item in selected_items:
                old_code = self.institutions_tree.item(item, 'values')[1]
                current_data = self.app.db_manager.get_institution_by_code(old_code)
                new_dom = ents["Email Domain:"].get()
                if new_dom: current_data['email'] = f"{current_data['institution_code']}{new_dom}"
                new_key = ents["Encryption Key:"].get()
                if new_key: current_data['encryption_key'] = new_key
                new_msg = ents["Message:"].get("1.0", tk.END).strip()
                if new_msg: current_data['message'] = new_msg
                self.app.db_manager.update_institution(old_code, current_data)
            self.load_institutions()
            win.destroy()
        ttk.Button(win, text="Update All", style="Multi.Glossy.TButton", command=apply_batch).grid(row=len(fields)+1, column=0, columnspan=2, pady=15)

    def delete_institution(self):
        sel = self.institutions_tree.selection()
        if not sel: return
        if messagebox.askyesno("Confirm", f"Delete {len(sel)} record(s)?"):
            for item in sel:
                self.app.db_manager.delete_institution(self.institutions_tree.item(item, 'values')[1])
            self.load_institutions()

    def setup_settings_tab(self, parent_frame):
        main_container = ttk.Frame(parent_frame, padding="10")
        main_container.pack(fill=tk.BOTH, expand=True)

        network_frame = ttk.LabelFrame(main_container, text="Network & Data (Shared Server)", padding="15")
        network_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        ttk.Label(network_frame, text="Shared Database Path (file_monitor.db):", font=("Helvetica", 9, "bold")).pack(anchor=tk.W, pady=(0, 2))
        db_frame = ttk.Frame(network_frame)
        db_frame.pack(fill=tk.X, pady=(0, 15))
        current_db = self.app.local_config.get('PATHS', 'db_path', fallback='') if hasattr(self.app, 'local_config') else ''
        self.db_path_var = tk.StringVar(value=current_db)
        ttk.Entry(db_frame, textvariable=self.db_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(db_frame, text="Browse", style="Glossy.TButton", command=self.browse_db).pack(side=tk.RIGHT)

        ttk.Label(network_frame, text="Master Config Path (master_config.ini):", font=("Helvetica", 9, "bold")).pack(anchor=tk.W, pady=(0, 2))
        cfg_frame = ttk.Frame(network_frame)
        cfg_frame.pack(fill=tk.X, pady=(0, 15))
        current_cfg = self.app.local_config.get('PATHS', 'master_config_path', fallback='') if hasattr(self.app, 'local_config') else ''
        self.config_path_var = tk.StringVar(value=current_cfg)
        ttk.Entry(cfg_frame, textvariable=self.config_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(cfg_frame, text="Browse", style="Glossy.TButton", command=self.browse_config).pack(side=tk.RIGHT)

        self.status_lbl = ttk.Label(network_frame, text="Status: Ready", foreground="gray", font=("Helvetica", 10, "italic"))
        self.status_lbl.pack(anchor=tk.W, pady=(10, 0))

        prefs_frame = ttk.LabelFrame(main_container, text="Local User Preferences", padding="15")
        prefs_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        ttk.Label(prefs_frame, text="After Successful Draft Creation:", font=("Helvetica", 9, "bold")).pack(anchor=tk.W, pady=(0, 5))
        current_post = self.app.local_config.get('PREFS', 'post_process', fallback='keep') if hasattr(self.app, 'local_config') else 'keep'
        self.post_process_var = tk.StringVar(value=current_post)
        
        ttk.Radiobutton(prefs_frame, text="Keep original files in 'To Send'", variable=self.post_process_var, value='keep').pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(prefs_frame, text="Move original files to Master Archive", variable=self.post_process_var, value='archive').pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(prefs_frame, text="Delete original files permanently", variable=self.post_process_var, value='delete').pack(anchor=tk.W, pady=2)

        ttk.Label(prefs_frame, text="Max Email Attachment Limit (MB):", font=("Helvetica", 9, "bold")).pack(anchor=tk.W, pady=(20, 5))
        current_max = self.app.local_config.get('PREFS', 'max_size_mb', fallback='10.0') if hasattr(self.app, 'local_config') else '10.0'
        
        self.max_size_var = tk.StringVar(value=str(current_max))
        ttk.Entry(prefs_frame, textvariable=self.max_size_var, width=15).pack(anchor=tk.W)

        bottom_frame = ttk.Frame(parent_frame)
        bottom_frame.pack(fill=tk.X, pady=10, padx=20)
        ttk.Button(bottom_frame, text="Save & Reconnect", style="Add.Glossy.TButton", command=self.save_settings).pack(side=tk.RIGHT)

    def browse_db(self):
        path = filedialog.askopenfilename(title="Select Shared Database", filetypes=[("SQLite DB", "*.db"), ("All Files", "*.*")])
        if path:
            self.db_path_var.set(path)

    def browse_config(self):
        path = filedialog.askopenfilename(title="Select Master Config", filetypes=[("INI Files", "*.ini"), ("All Files", "*.*")])
        if path:
            self.config_path_var.set(path)

    def save_settings(self):
        settings_data = {
            'db_path': self.db_path_var.get(),
            'master_config_path': self.config_path_var.get(),
            'post_process': self.post_process_var.get(),
            'max_size_mb': self.max_size_var.get()
        }
        success = self.app.save_local_settings(settings_data) 
        
        if success:
            self.status_lbl.config(text="Status: Connected to Master Storage", foreground="green")
            messagebox.showinfo("Settings Saved", "Local settings updated and reconnected successfully.")
        else:
            self.status_lbl.config(text="Status: Connection Failed. Check Paths.", foreground="red")

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path_entry.delete(0, tk.END)
            self.folder_path_entry.insert(0, folder)

    def start_monitoring(self):
        path = self.folder_path_entry.get()
        if path and self.app.start_monitoring(path):
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)

    def stop_monitoring(self):
        self.app.stop_monitoring()
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

    def log_activity(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def load_institutions(self):
        if not self.app.db_manager: return
        for i in self.institutions_tree.get_children(): self.institutions_tree.delete(i)
        for inst in self.app.db_manager.get_all_institutions(): self.institutions_tree.insert("", "end", values=inst)

    def add_batch(self, bd):
        is_warning_panel = not bd.get('is_today', True)
        
        panel_style = "Warning.TLabelframe" if is_warning_panel else "Normal.TLabelframe"
        cb_style = "Warning.TCheckbutton" if is_warning_panel else "Normal.TCheckbutton"
        warning_lbl_style = "WarningText.TLabel" if is_warning_panel else "NormalText.TLabel"
        btn_frame_style = "Warning.TFrame" if is_warning_panel else "Normal.TFrame"

        panel = ttk.LabelFrame(self.batches_inner_frame, text=f"Batch {bd['batch_number']} ({bd['institution_code']})", padding="10", style=panel_style)
        panel.pack(fill=tk.X, padx=5, pady=5)
        bd['file_vars'] = {}
        
        requires_myob = bd['institution_code'] in ['NAB', 'NABC', 'CCA']
        myob_found = False

        for file in bd['files']:
            var = tk.BooleanVar(value=True)
            bd['file_vars'][file] = var
            ttk.Checkbutton(panel, text=os.path.basename(file), variable=var, style=cb_style).pack(anchor='w', padx=15)
            
            if requires_myob and file.lower().endswith(('.xls', '.xlsx', '.csv')):
                from encryption_utils import has_myob_id
                if has_myob_id(file):
                    myob_found = True

        if requires_myob and not myob_found:
            ttk.Label(panel, text="⚠ Cannot Find MYOB_ID. Please check if the correct file is selected.", style=warning_lbl_style, wraplength=280).pack(anchor='w', padx=15, pady=5)
            bd['missing_myob'] = True

        btn_frame = ttk.Frame(panel, style=btn_frame_style)
        btn_frame.pack(pady=5)
        
        ttk.Button(btn_frame, text="Create Draft", style="Edit.Glossy.TButton", command=lambda b=bd: self.confirm_and_draft(b)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Close Without Drafting", style="Delete.Glossy.TButton", command=lambda bn=bd['batch_number']: self.cancel_batch_processing(bn)).pack(side=tk.LEFT, padx=5)
        
        self.batch_panels[bd['batch_number']] = panel
        self.batches_canvas.config(scrollregion=self.batches_canvas.bbox("all"))

    def confirm_and_draft(self, bd):
        if not bd.get('is_today', True):
            if not messagebox.askyesno("Date Warning", "The date of the file is not today.\nAre you sure you want to send this?"):
                self.cancel_batch_processing(bd['batch_number'])
                return
                
        if bd.get('missing_myob', False):
            if not messagebox.askyesno("Missing MYOB_ID", "Cannot Find MYOB_ID.\nPlease check if correct file is selected.\n\nProceed anyway?"):
                return
                
        self.app.process_batch(bd)

    def cancel_batch_processing(self, batch_number):
        if messagebox.askyesno("Cancel", "Discard this detected batch without creating drafts?"):
            self.remove_batch_panel(batch_number)
            self.log_activity(f"Batch {batch_number} processing cancelled by user.")

    def remove_batch_panel(self, bn):
        if bn in self.batch_panels:
            self.batch_panels[bn].destroy()
            del self.batch_panels[bn]

    def load_processed_batches(self):
        if not self.app.db_manager: return
        for i in self.processed_tree.get_children(): self.processed_tree.delete(i)
        for b in self.app.db_manager.get_sent_emails(): self.processed_tree.insert("", "end", values=b)
        self.autofit_columns()
    
    def add_processed_batch(self, bd): 
        self.load_processed_batches()
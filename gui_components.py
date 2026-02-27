import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font
import os
from datetime import datetime

class MainWindow:
    def __init__(self, root, app_controller):
        self.root = root
        self.app = app_controller
        
        # 1. Startup User Selection (Before building the main UI)
        self.current_user = self.select_user_startup()
        
        self.root.title(f"V.O.I.D. - Verified On-boarding Institutional Dispatcher [User: {self.current_user}]")
        
        self.setup_styles()
        
        # Main Layout Container
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tab 1: Monitoring Dashboard
        self.main_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.main_tab, text="Monitoring")
        self.setup_monitoring_tab(self.main_tab)
        
        # Tab 2: Institution Database Management
        self.institutions_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.institutions_tab, text="Institutions")
        self.setup_institutions_tab(self.institutions_tab)
        
        self.batch_panels = {}
        self.load_institutions()
        self.load_processed_batches()

    def select_user_startup(self):
        """Pops up a selection window for the dispatcher name at startup."""
        staff_list = self.app.config.get('Users', 'staff_list', fallback="User").split(',')
        staff_list = [s.strip() for s in staff_list]
        
        # Modal selection window
        win = tk.Toplevel(self.root)
        win.title("V.O.I.D. Login")
        win.geometry("350x180")
        win.resizable(False, False)
        win.grab_set()  # Focus user on this window
        
        # Center the window
        win.update_idletasks()
        x = (win.winfo_screenwidth() // 2) - (win.winfo_width() // 2)
        y = (win.winfo_screenheight() // 2) - (win.winfo_height() // 2)
        win.geometry(f'+{x}+{y}')

        ttk.Label(win, text="Who is dispatching today?", font=("Helvetica", 11, "bold")).pack(pady=15)
        
        user_var = tk.StringVar(value=staff_list[0])
        combo = ttk.Combobox(win, textvariable=user_var, values=staff_list, state="readonly", font=("Helvetica", 10))
        combo.pack(pady=5, padx=30, fill=tk.X)
        
        selected_user = [staff_list[0]] 
        
        def on_confirm():
            selected_user[0] = user_var.get()
            win.destroy()
            
        btn = ttk.Button(win, text="Start System", command=on_confirm)
        btn.pack(pady=20)
        
        # Prevent closing without selection
        win.protocol("WM_DELETE_WINDOW", lambda: self.root.destroy())
        self.root.wait_window(win)
        return selected_user[0]

    def setup_styles(self):
        """Sets up the V.O.I.D. signature purple and gold theme."""
        self.style = ttk.Style()
        self.style.theme_use("alt")
        purple, gold = "#4B0082", "#FFD700"
        
        self.style.configure("TNotebook", background=purple)
        self.style.configure("TNotebook.Tab", padding=[15, 5], background=gold, foreground="black", font=("Helvetica", 10, "bold"))
        self.style.map("TNotebook.Tab", background=[("selected", "#FFA500")])
        self.style.configure("TFrame", background="#F0F0F0")
        self.style.configure("TLabelFrame", background="#F0F0F0", foreground=purple, font=("Helvetica", 11, "bold"))
        self.style.configure("Treeview", rowheight=28, font=("Helvetica", 9))
        self.style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"))
        self.style.map("Treeview", background=[("selected", "#3498db")], foreground=[("selected", "white")])

    def setup_monitoring_tab(self, parent_frame):
        """Builds the main dashboard for folder watching and history."""
        # Top: Folder Selection
        self.folder_frame = ttk.LabelFrame(parent_frame, text="Folder Monitoring", padding="10")
        self.folder_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.folder_path_entry = ttk.Entry(self.folder_frame, width=50)
        self.folder_path_entry.grid(row=0, column=0, padx=5, pady=5)
        ttk.Button(self.folder_frame, text="Browse", command=self.browse_folder).grid(row=0, column=1, padx=5, pady=5)
        
        self.start_button = ttk.Button(self.folder_frame, text="Start Monitor", command=self.start_monitoring)
        self.start_button.grid(row=0, column=2, padx=5, pady=5)
        self.stop_button = ttk.Button(self.folder_frame, text="Stop", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Button(self.folder_frame, text="Purge Folder", command=self.app.delete_all_files).grid(row=0, column=4, padx=5, pady=5)

        # Split View: Detected vs Processed
        self.batches_main_frame = ttk.Frame(parent_frame)
        self.batches_main_frame.pack(fill=tk.BOTH, expand=True)

        # Left: Detected Batches
        self.detected_batches_frame = ttk.LabelFrame(self.batches_main_frame, text="Live Inbox (Detected)", padding="10")
        self.detected_batches_frame.pack(fill=tk.BOTH, expand=True, padx=(0, 10), side=tk.LEFT)
        
        self.batches_canvas = tk.Canvas(self.detected_batches_frame, bg="#E8E8E8")
        self.batches_scrollbar = ttk.Scrollbar(self.detected_batches_frame, orient="vertical", command=self.batches_canvas.yview)
        self.batches_inner_frame = ttk.Frame(self.batches_canvas)
        self.batches_inner_frame.bind("<Configure>", lambda e: self.batches_canvas.configure(scrollregion=self.batches_canvas.bbox("all")))
        self.batches_canvas.create_window((0, 0), window=self.batches_inner_frame, anchor="nw")
        self.batches_canvas.configure(yscrollcommand=self.batches_scrollbar.set)
        self.batches_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.batches_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Right: Processed Batches (Shared History)
        self.processed_batches_frame = ttk.LabelFrame(self.batches_main_frame, text="Shared Dispatch History", padding="10")
        self.processed_batches_frame.pack(fill=tk.BOTH, expand=True, side=tk.RIGHT)

        # Triple-Filter Search Bar
        self.filter_frame = ttk.Frame(self.processed_batches_frame)
        self.filter_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(self.filter_frame, text="Filter (Batch, Code, or DD/MM/YYYY):", font=("Helvetica", 9, "italic")).pack(side=tk.LEFT, padx=2)
        self.search_entry = ttk.Entry(self.filter_frame)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.search_entry.bind("<KeyRelease>", lambda e: self.filter_history())

        self.processed_tree = ttk.Treeview(self.processed_batches_frame, 
                                          columns=("inst", "batch", "date", "time", "files", "user", "recs"), 
                                          show="headings")
        
        heads = [("inst", "Inst"), ("batch", "Batch #"), ("date", "Date"), ("time", "Time"), 
                 ("files", "Files Sent"), ("user", "Dispatcher"), ("recs", "Records")]
        
        for col, head in heads:
            self.processed_tree.heading(col, text=head)
            self.processed_tree.column(col, anchor=tk.CENTER)

        self.processed_tree.pack(fill=tk.BOTH, expand=True)
        
        # Bottom: Activity Log
        self.log_frame = ttk.LabelFrame(parent_frame, text="System Log", padding="10")
        self.log_frame.pack(fill=tk.X, pady=(10, 0))
        self.log_text = tk.Text(self.log_frame, state="disabled", wrap=tk.WORD, height=4, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def autofit_columns(self):
        """Calculates column widths based on content + 2 characters of padding."""
        padding = 20 # Approx 2 characters in pixels
        for col in self.processed_tree["columns"]:
            # Measure header
            max_w = font.Font().measure(self.processed_tree.heading(col)['text']) + padding
            # Measure all rows for this column
            for item in self.processed_tree.get_children():
                cell_val = str(self.processed_tree.set(item, col))
                cell_w = font.Font().measure(cell_val) + padding
                if cell_w > max_w:
                    max_w = cell_w
            self.processed_tree.column(col, width=max_w)

    def load_processed_batches(self):
        """Refreshes the table from the shared DB and applies autofit."""
        for i in self.processed_tree.get_children(): self.processed_tree.delete(i)
        batches = self.app.db_manager.get_sent_emails()
        for b in batches:
            self.processed_tree.insert("", "end", values=b)
        self.autofit_columns()

    def filter_history(self):
        """Filters search results across Batch ID, Institution, and Standardized Date."""
        term = self.search_entry.get()
        for i in self.processed_tree.get_children(): self.processed_tree.delete(i)
        results = self.app.db_manager.search_sent_emails(term)
        for b in results:
            self.processed_tree.insert("", "end", values=b)
        self.autofit_columns()

    def add_batch(self, bd):
        """Creates a dynamic panel for a detected batch with Create Draft button."""
        panel = ttk.LabelFrame(self.batches_inner_frame, text=f"Batch {bd['batch_number']} - {bd['institution_code']}", padding="10")
        panel.pack(fill=tk.X, padx=5, pady=5)
        
        bd['file_vars'] = {}
        for file in bd['files']:
            var = tk.BooleanVar(value=True)
            bd['file_vars'][file] = var
            ttk.Checkbutton(panel, text=os.path.basename(file), variable=var).pack(anchor='w', padx=10)
            
        ttk.Button(panel, text="奪 Create Draft", command=lambda b=bd: self.app.process_batch(b)).pack(pady=8)
        self.batch_panels[bd['batch_number']] = panel
        self.batches_canvas.config(scrollregion=self.batches_canvas.bbox("all"))

    # --- Standard GUI Helpers ---
    def setup_institutions_tab(self, parent_frame):
        frame = ttk.LabelFrame(parent_frame, text="Database Management", padding="10")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.institutions_tree = ttk.Treeview(frame, columns=("c_code", "i_code", "email", "key", "msg"), show="headings")
        for col, text in zip(self.institutions_tree["columns"], ["County", "Inst Code", "Email", "Password", "Message"]):
            self.institutions_tree.heading(col, text=text)
        self.institutions_tree.pack(fill=tk.BOTH, expand=True)
        # (Buttons for Add/Edit/Delete would go here as in previous versions)

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
        for i in self.institutions_tree.get_children(): self.institutions_tree.delete(i)
        for inst in self.app.db_manager.get_all_institutions(): self.institutions_tree.insert("", "end", values=inst)

    def remove_batch_panel(self, bn):
        if bn in self.batch_panels:
            self.batch_panels[bn].destroy()
            del self.batch_panels[bn]
            self.batches_canvas.config(scrollregion=self.batches_canvas.bbox("all"))
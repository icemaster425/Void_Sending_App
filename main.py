import tkinter as tk
from tkinter import messagebox
import os
import sys
import queue
import logging
from datetime import datetime
from configparser import ConfigParser

# Importing custom V.O.I.D. modules
from gui_components import MainWindow
from file_monitor import FileMonitor
from database_manager import DatabaseManager
from outlook_integration import OutlookIntegration
from encryption_utils import (
    zip_files_with_password, 
    transform_excel, 
    check_pdf_integrity, 
    split_pdf_pages
)

class FileMonitorApp:
    def __init__(self):
        self.root = tk.Tk()
        # Handle pathing for script and EXE environments [cite: 1, 2]
        self.base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        
        # 1. Load Configuration [cite: 1, 3]
        self.config = self.load_config()
        
        # 2. Setup Shared Database 
        # Uses the network path defined in SharedSettings
        db_path = self.config.get('SharedSettings', 'db_location', fallback='file_monitor.db')
        self.db_manager = DatabaseManager(db_path)
        
        # 3. Setup Folders and Systems [cite: 1, 7, 10]
        self.default_monitor_path = os.path.join(self.base_dir, "To Send")
        if not os.path.exists(self.default_monitor_path):
            os.makedirs(self.default_monitor_path)
            
        self.setup_logging()
        self.outlook_integration = OutlookIntegration()
        self.message_queue = queue.Queue()
        self.monitoring = False
        
        # 4. Initialize UI (Triggers User Selection at Startup) 
        self.gui = MainWindow(self.root, self)
        self.gui.folder_path_entry.insert(0, self.default_monitor_path)
        
        self.process_messages()

    def load_config(self):
        """Loads config.ini with support for multi-user and recipes."""
        config_path = os.path.join(self.base_dir, 'config.ini')
        config = ConfigParser()
        if os.path.exists(config_path):
            config.read(config_path)
        return config

    def setup_logging(self):
        log_level_str = self.config.get('SYSTEM', 'log_level', fallback='INFO').upper()
        level = getattr(logging, log_level_str, logging.INFO)
        log_path = os.path.join(self.base_dir, 'file_monitor.log')
        logging.basicConfig(level=level, format='%(asctime)s - %(levelname)s - %(message)s', 
                            handlers=[logging.FileHandler(log_path), logging.StreamHandler(sys.stdout)])

    def process_batch(self, batch_data):
        """Final coordination of recipes, size guards, and dispatch."""
        ic = batch_data.get('institution_code')
        bn = batch_data.get('batch_number')
        raw_files = [f for f, var in batch_data['file_vars'].items() if var.get()]
        info = self.db_manager.get_institution_by_code(ic)
        
        if not info:
            messagebox.showwarning("Warning", f"Institution Code '{ic}' not found.")
            return

        # 1. Look up Institution Recipe
        recipe = self.config.get('INSTITUTION_RECIPES', ic, fallback='normal')
        temp_files = []
        record_count = 0

        try:
            processed_files = []
            for f in raw_files:
                ext = os.path.splitext(f)[1].lower()
                
                # Apply PDF Integrity Check
                if ext == '.pdf':
                    healthy, msg = check_pdf_integrity(f)
                    if not healthy:
                        raise Exception(f"File {os.path.basename(f)} failed check: {msg}")
                    
                    # Apply PDF Splitting if recipe requires
                    if recipe == 'split_pdf':
                        splits = split_pdf_pages(f, self.base_dir)
                        processed_files.extend(splits)
                        temp_files.extend(splits)
                        continue

                # Apply Excel Transformations (BSB Split or Version Conversion)
                if ext in ['.xls', '.xlsx']:
                    new_path, count = transform_excel(f, recipe, self.base_dir)
                    processed_files.append(new_path)
                    temp_files.append(new_path)
                    record_count = count
                    continue
                
                processed_files.append(f)

            # 2. Encrypt and Validate Size
            zip_name = f"{ic}_{bn}.zip"
            zip_path = os.path.join(self.base_dir, zip_name)
            zip_files_with_password(processed_files, zip_path, info['encryption_key'])
            
            # STRICT 10MB GUARD
            file_size_mb = os.path.getsize(zip_path) / (1024 * 1024)
            limit = self.config.getfloat('Settings', 'max_email_size', fallback=10.0)
            
            if file_size_mb > limit:
                os.remove(zip_path)
                messagebox.showerror("Limit Exceeded", 
                    f"Batch is {file_size_mb:.2f}MB (Limit: {limit}MB).\n"
                    "Draft not created and not logged to database.")
                return

            # 3. Format Subject and Signature
            date_str = datetime.now().strftime('%d/%m/%Y')
            subject = self.config.get('EMAIL', 'subject_template').format(
                inst_code=ic, date=date_str, batch_number=bn, count=record_count
            )
            
            # Build Signature with selected Dispatcher
            footer = self.config.get('EMAIL', 'email_footer', fallback="")
            final_body = f"{info['message']}\n\n{footer} {self.gui.current_user}"

            # 4. Create Draft and Log to Shared DB
            success, msg = self.outlook_integration.create_draft(info['email'], subject, final_body, [zip_path])
            
            if success:
                # Logs include the standardized date and user name
                self.db_manager.add_sent_email(ic, bn, info['email'], subject, "ZIP", 
                                             processed_files, self.gui.current_user, record_count)
                self.gui.load_processed_batches()
                self.gui.remove_batch_panel(bn)
                self.file_monitor.remove_from_queue(bn)
                self.gui.log_activity(f"SUCCESS: {ic} Batch {bn} by {self.gui.current_user}")
                
        except Exception as e:
            messagebox.showerror("Process Error", str(e))
        finally:
            # Cleanup temp zips and converted files
            if 'zip_path' in locals() and os.path.exists(zip_path): os.remove(zip_path)
            for tf in temp_files:
                if os.path.exists(tf): os.remove(tf)

    def start_monitoring(self, folder_path):
        self.file_monitor = FileMonitor(folder_path, self.db_manager, self.message_queue, 
                                        self.file_extensions, self.batch_length)
        self.file_monitor.start()
        self.monitoring = True
        return True

    def stop_monitoring(self):
        if self.file_monitor: self.file_monitor.stop()
        self.monitoring = False

    def process_messages(self):
        while not self.message_queue.empty():
            msg = self.message_queue.get()
            if msg['type'] == 'activity': self.gui.log_activity(msg['data'])
            elif msg['type'] == 'batch_detected': self.gui.add_batch(msg['data'])
        self.root.after(100, self.process_messages)

    def delete_all_files(self):
        folder = self.gui.folder_path_entry.get()
        if messagebox.askyesno("System", f"Purge all files in {folder}?"):
            for f in os.listdir(folder): os.remove(os.path.join(folder, f))

    def run(self): self.root.mainloop()

if __name__ == '__main__':
    app = FileMonitorApp()
    app.run()
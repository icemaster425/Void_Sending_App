import tkinter as tk
from tkinter import messagebox
import os
import sys
import queue
import logging
import shutil
from datetime import datetime
from configparser import ConfigParser

# Local imports
from gui_components import MainWindow
from file_monitor import FileMonitor
from database_manager import DatabaseManager
from outlook_integration import OutlookIntegration

# Expanded imports for recipes and security
from encryption_utils import (
    zip_files_with_password, 
    transform_excel, 
    check_file_integrity, 
    split_pdf_pages,
    convert_pdf_to_tiff
)

class FileMonitorApp:
    def __init__(self):
        self.root = tk.Tk()
        self.base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        
        self.setup_logging()
        
        # 1. Load Local Config (Settings Tab preferences)
        self.local_config_path = os.path.join(self.base_dir, 'local_config.ini')
        self.local_config = self.load_local_config()
        
        # 2. Setup Default Folder Logic
        self.default_monitor_path = os.path.join(self.base_dir, "To Send")
        if not os.path.exists(self.default_monitor_path):
            os.makedirs(self.default_monitor_path)
            
        self.root.title("V.O.I.D. - Initializing...")
        self.root.geometry("1200x800")
        
        # 3. Initialize Shared Resources (Will be None if paths aren't set yet)
        self.db_manager = None
        self.master_config = None
        self.outlook_integration = OutlookIntegration()
        
        self.file_monitor = None
        self.monitoring = False
        self.message_queue = queue.Queue()
        
        # Load the network resources if the local config has them
        self.connect_to_master()
        
        # Initialize UI
        self.gui = MainWindow(self.root, self)
        self.gui.folder_path_entry.insert(0, self.default_monitor_path)
        
        # Update UI connection status if successful on startup
        if self.db_manager and self.master_config:
            self.gui.status_lbl.config(text="Status: Connected to Master Storage", foreground="green")
            self.gui.load_institutions()
            self.gui.load_processed_batches()
        
        self.process_messages()

    def setup_logging(self):
        log_path = os.path.join(self.base_dir, 'file_monitor.log')
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', 
                            handlers=[logging.FileHandler(log_path), logging.StreamHandler(sys.stdout)])

    def load_local_config(self):
        config = ConfigParser()
        if os.path.exists(self.local_config_path):
            config.read(self.local_config_path)
        else:
            config['PATHS'] = {'db_path': '', 'master_config_path': ''}
            config['PREFS'] = {'post_process': 'keep', 'max_size_mb': '10.0'}
            with open(self.local_config_path, 'w') as f:
                config.write(f)
        return config

    def save_local_settings(self, settings_data):
        """Called by the UI to save network paths and user preferences."""
        if not self.local_config.has_section('PATHS'):
            self.local_config.add_section('PATHS')
        if not self.local_config.has_section('PREFS'):
            self.local_config.add_section('PREFS')
            
        self.local_config.set('PATHS', 'db_path', settings_data['db_path'])
        self.local_config.set('PATHS', 'master_config_path', settings_data['master_config_path'])
        self.local_config.set('PREFS', 'post_process', settings_data['post_process'])
        self.local_config.set('PREFS', 'max_size_mb', settings_data['max_size_mb'])
        
        with open(self.local_config_path, 'w') as f:
            self.local_config.write(f)
            
        return self.connect_to_master()

    def connect_to_master(self):
        """Attempts to load the Master Database and Master Config from the S: drive."""
        db_path = self.local_config.get('PATHS', 'db_path', fallback='')
        cfg_path = self.local_config.get('PATHS', 'master_config_path', fallback='')
        
        if not db_path or not cfg_path or not os.path.exists(db_path) or not os.path.exists(cfg_path):
            return False
            
        try:
            self.db_manager = DatabaseManager(db_path)
            self.master_config = ConfigParser()
            self.master_config.read(cfg_path)
            
            # Read dynamic settings from Master
            self.file_extensions = self.master_config.get('MONITORING', 'file_extensions', fallback='.xlsx,.xls,.pdf').split(',')
            self.batch_length = self.master_config.get('MONITORING', 'batch_length', fallback='6')
            
            # Update UI Trees if the GUI is already loaded
            if hasattr(self, 'gui'):
                self.gui.load_institutions()
                self.gui.load_processed_batches()
                
            return True
        except Exception as e:
            logging.error(f"Failed to connect to master storage: {e}")
            return False

    def delete_all_files(self):
        folder = self.gui.folder_path_entry.get()
        if not folder or not os.path.exists(folder): return

        if messagebox.askyesno("System", f"Confirm deletion of all files in:\n{folder}?"):
            try:
                count = 0
                for filename in os.listdir(folder):
                    file_path = os.path.join(folder, filename)
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                        count += 1
                self.gui.log_activity(f"Action: Purged {count} files.")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    def process_batch(self, batch_data):
        if not self.db_manager or not self.master_config:
            messagebox.showerror("Network Error", "Not connected to Master Storage. Please check your Settings tab.")
            return

        ic = batch_data.get('institution_code')
        bn = batch_data.get('batch_number')
        raw_files = [f for f, var in batch_data['file_vars'].items() if var.get()]
        info = self.db_manager.get_institution_by_code(ic)
        
        if not info:
            messagebox.showwarning("Warning", f"New Code Detected: [{ic}].\nPlease assign an email and password in the Institutions tab before dispatching.")
            return

        recipe_string = self.master_config.get('INSTITUTION_RECIPES', ic, fallback='')
        recipes = [r.strip() for r in recipe_string.split(',')] if recipe_string else []
        
        processed_files = []
        temp_files = []
        record_count = 0

        try:
            # STEP 0: Integrity Check
            for f in raw_files:
                healthy, msg = check_file_integrity(f)
                if not healthy:
                    raise Exception(f"Integrity Error in {os.path.basename(f)}: {msg}")

            # STEP 1 & 2: Apply Transformations & Counting Recipes
            for f in raw_files:
                ext = os.path.splitext(f)[1].lower()
                
                if ext == '.pdf':
                    if 'split_pdf' in recipes:
                        splits = split_pdf_pages(f, self.base_dir)
                        processed_files.extend(splits)
                        temp_files.extend(splits)
                        continue
                    elif 'pdf_to_tiff' in recipes:
                        tiffs = convert_pdf_to_tiff(f, self.base_dir)
                        processed_files.extend(tiffs)
                        temp_files.extend(tiffs)
                        continue

                if ext in ['.xls', '.xlsx']:
                    # Passes the recipes to the engine to handle BSB splits and format swaps
                    new_path, rows = transform_excel(f, recipes, self.base_dir)
                    processed_files.append(new_path)
                    temp_files.append(new_path)
                    
                    if 'add_count' in recipes:
                        record_count = rows
                    continue
                
                # If no specific recipe applied, just pass the original file through
                processed_files.append(f)

            # STEP 3: The Vault (Encryption)
            zip_path = os.path.join(self.base_dir, f"{ic}_{bn}.zip")
            zip_files_with_password(processed_files, zip_path, info['encryption_key'], f"{ic}_{bn}")
            
            # STEP 4: The Strict 10MB Guard
            file_size_mb = os.path.getsize(zip_path) / (1024 * 1024)
            limit = self.local_config.getfloat('PREFS', 'max_size_mb', fallback=10.0)
            
            if file_size_mb > limit:
                raise Exception(f"Batch is {file_size_mb:.2f}MB (Limit: {limit}MB).\nDraft aborted. Please use a split recipe.")

            # STEP 5: Dispatch Assembly & Logging
            subject_template = self.master_config.get('EMAIL_TEMPLATES', 'subject_template', fallback="{inst_code} Loads {date} Batch {batch_number}")
            footer = "\n\n" + self.master_config.get('EMAIL_TEMPLATES', 'email_footer', fallback="Regards,")
            core_msg = info['message'] if info.get('message') else ""
            
            final_body = f"{core_msg}{footer}\n{self.gui.current_user}"
            date_str = datetime.now().strftime('%d/%m/%Y')
            subject = subject_template.format(inst_code=ic, date=date_str, batch_number=bn)
            
            # Apply count to subject if recipe demanded it
            if 'add_count' in recipes and record_count > 0:
                subject += f" ({record_count})"

            success, msg = self.outlook_integration.create_draft(info['email'], subject, final_body, [zip_path])
            if success:
                self.db_manager.add_sent_email(ic, bn, info['email'], subject, "ZIP", processed_files, self.gui.current_user, record_count)
                self.gui.add_processed_batch(batch_data)
                self.gui.remove_batch_panel(bn)
                self.file_monitor.remove_from_queue(bn)
                self.gui.log_activity(f"SUCCESS: {ic} Batch {bn} by {self.gui.current_user}")
                
                # STEP 6: Post-Process Action
                post_action = self.local_config.get('PREFS', 'post_process', fallback='keep')
                if post_action == 'delete':
                    for f in raw_files:
                        os.remove(f)
                elif post_action == 'archive':
                    master_db_dir = os.path.dirname(self.local_config.get('PATHS', 'db_path'))
                    archive_folder = os.path.join(master_db_dir, '..', 'Archive', f"{date_str.replace('/','-')}_{ic}_{bn}")
                    os.makedirs(archive_folder, exist_ok=True)
                    for f in raw_files:
                        shutil.move(f, os.path.join(archive_folder, os.path.basename(f)))
                
        except Exception as e:
            messagebox.showerror("Process Error", str(e))
            self.gui.log_activity(f"ERROR: {str(e)}")
        finally:
            # Clean up generated ZIPs and temp Excel/PDF/TIFF files so they don't clutter the drive
            if 'zip_path' in locals() and os.path.exists(zip_path): 
                os.remove(zip_path)
            for tf in temp_files:
                if os.path.exists(tf): 
                    os.remove(tf)

    def start_monitoring(self, folder_path):
        if not self.db_manager:
            messagebox.showerror("Error", "Connect to Master Database in Settings first.")
            return False
            
        try:
            self.file_monitor = FileMonitor(folder_path, self.db_manager, self.message_queue, self.file_extensions, self.batch_length)
            self.file_monitor.start()
            self.monitoring = True
            return True
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return False

    def stop_monitoring(self):
        if self.file_monitor:
            self.file_monitor.stop()
            self.monitoring = False

    def process_messages(self):
        while not self.message_queue.empty():
            msg = self.message_queue.get()
            if msg['type'] == 'activity': self.gui.log_activity(msg['data'])
            elif msg['type'] == 'batch_detected': self.gui.add_batch(msg['data'])
        self.root.after(100, self.process_messages)

    def run(self): 
        self.root.mainloop()

if __name__ == '__main__':
    app = FileMonitorApp()
    app.run()
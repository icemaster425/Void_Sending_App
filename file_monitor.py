import os
import time
import threading
import re
from datetime import datetime

class FileMonitor(threading.Thread):
    def __init__(self, folder_path, db_manager, message_queue, file_extensions, batch_length):
        super().__init__()
        self.folder_path = folder_path
        self.db_manager = db_manager
        self.message_queue = message_queue
        
        # Ensures all extensions from config are lowercase and stripped of spaces
        self.file_extensions = [ext.strip().lower() for ext in file_extensions]
        
        # Ensures batch length is treated as an integer
        self.batch_length = int(batch_length) if str(batch_length).isdigit() else 6
        
        self.running = False
        self.detected_batches = set()

    def run(self):
        self.running = True
        self.message_queue.put({'type': 'activity', 'data': f"Started monitoring: {self.folder_path}"})
        
        # Background loop runs every 3 seconds while active
        while self.running:
            self.scan_folder()
            time.sleep(3)

    def scan_folder(self):
        if not os.path.exists(self.folder_path):
            return

        current_batches = {}
        today_str = datetime.now().strftime('%d%m%Y')

        # STEP 1: Scan all files and group them by Batch Number
        for filename in os.listdir(self.folder_path):
            file_path = os.path.join(self.folder_path, filename)
            
            # Skip folders
            if not os.path.isfile(file_path):
                continue

            # Check if file matches our allowed extensions
            ext = os.path.splitext(filename)[1].lower()
            if ext not in self.file_extensions:
                continue

            inst_code = "PENDING"
            batch_num = None
            is_today = True # Default for Excel which doesn't have dates in name

            # RULE 1: Excel Files (BatchNumber_InstitutionCode.ext)
            if ext in ['.xls', '.xlsx']:
                match = re.match(r'^(\d+)_([A-Za-z0-9]+)', filename)
                if match:
                    batch_num = match.group(1)
                    inst_code = match.group(2).upper()

            # RULE 2: PDF Files (Date_BatchNumber.pdf)
            elif ext == '.pdf':
                match = re.match(r'^(\d+)_(\d+)', filename)
                if match:
                    file_date = match.group(1)
                    batch_num = match.group(2)
                    is_today = (file_date == today_str)
                    inst_code = "PENDING"

            # Grouping Logic: Add the file to the correct Batch dictionary
            if batch_num:
                if batch_num not in current_batches:
                    current_batches[batch_num] = {
                        'institution_code': inst_code,
                        'batch_number': batch_num,
                        'files': [],
                        'is_today': True
                    }
                
                current_batches[batch_num]['files'].append(file_path)
                
                # Flag the entire batch if even one file is outdated
                if not is_today:
                    current_batches[batch_num]['is_today'] = False
                
                # Update the whole group so PDFs know who they belong to
                if inst_code != "PENDING":
                    current_batches[batch_num]['institution_code'] = inst_code

        # STEP 2: Push completed batches to the GUI
        for b_num, data in current_batches.items():
            if data['institution_code'] == "PENDING":
                continue

            if b_num not in self.detected_batches:
                self.detected_batches.add(b_num)
                
                self.message_queue.put({
                    'type': 'batch_detected',
                    'data': data
                })
                
                file_count = len(data['files'])
                self.message_queue.put({
                    'type': 'activity',
                    'data': f"Detected: {data['institution_code']} Batch {b_num} ({file_count} files)"
                })

    def stop(self):
        self.running = False
        self.message_queue.put({'type': 'activity', 'data': "Monitoring stopped."})

    def remove_from_queue(self, batch_number):
        pass
import os
import time
import threading
import re

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
        
        # Keeps track of batches already pushed to the UI so it doesn't spam duplicates
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

        # STEP 1: Scan all files and group them by Batch Number
        for filename in os.listdir(self.folder_path):
            file_path = os.path.join(self.folder_path, filename)
            
            # Skip folders
            if not os.path.isfile(file_path):
                continue

            # Check if file matches our allowed extensions (.pdf, .xls, .xlsx)
            ext = os.path.splitext(filename)[1].lower()
            if ext not in self.file_extensions:
                continue

            inst_code = "PENDING"
            batch_num = None

            # RULE 1: Excel Files (BatchNumber_InstitutionCode.ext)
            # Example: 220593_ANZ.xlsx -> Batch: 220593, Code: ANZ
            if ext in ['.xls', '.xlsx']:
                match = re.match(r'^(\d+)_([A-Za-z0-9]+)', filename)
                if match:
                    batch_num = match.group(1)
                    inst_code = match.group(2).upper()

            # RULE 2: PDF Files (Date_BatchNumber.pdf)
            # Example: 20052025_220593.pdf -> Date: 20052025, Batch: 220593
            elif ext == '.pdf':
                match = re.match(r'^(\d+)_(\d+)', filename)
                if match:
                    # Group 2 is the batch number for PDFs based on your naming convention
                    batch_num = match.group(2)
                    inst_code = "PENDING"

            # Grouping Logic: Add the file to the correct Batch dictionary
            if batch_num:
                if batch_num not in current_batches:
                    current_batches[batch_num] = {
                        'institution_code': inst_code,
                        'batch_number': batch_num,
                        'files': []
                    }
                
                current_batches[batch_num]['files'].append(file_path)
                
                # If this file gave us the real Institution Code (from an Excel file), 
                # update the whole group so any PDFs in this batch know who they belong to.
                if inst_code != "PENDING":
                    current_batches[batch_num]['institution_code'] = inst_code

        # STEP 2: Push completed batches to the GUI
        for b_num, data in current_batches.items():
            # We cannot process a batch if we don't know the Institution Code yet
            # (e.g., if only the PDF has arrived in the folder so far)
            if data['institution_code'] == "PENDING":
                continue

            if b_num not in self.detected_batches:
                self.detected_batches.add(b_num)
                
                # Send the batch data to populate the UI panel
                self.message_queue.put({
                    'type': 'batch_detected',
                    'data': data
                })
                
                # Log the detection in the activity feed
                file_count = len(data['files'])
                self.message_queue.put({
                    'type': 'activity',
                    'data': f"Detected: {data['institution_code']} Batch {b_num} ({file_count} files)"
                })

    def stop(self):
        self.running = False
        self.message_queue.put({'type': 'activity', 'data': "Monitoring stopped."})

    def remove_from_queue(self, batch_number):
        """
        Called by main.py when a batch is successfully processed.
        We leave it in the detected_batches set to prevent the monitor from 
        re-adding it if the user has 'Keep Files' selected in their Settings tab.
        """
        pass
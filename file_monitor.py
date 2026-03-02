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
        
        # Ensures all extensions from config are lowercase and stripped of spaces for accurate matching
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

        # Scan all files in the directory
        for filename in os.listdir(self.folder_path):
            file_path = os.path.join(self.folder_path, filename)
            
            # Skip folders
            if not os.path.isfile(file_path):
                continue

            # Check if file matches our allowed extensions (.pdf, .xls, .xlsx)
            ext = os.path.splitext(filename)[1].lower()
            if ext not in self.file_extensions:
                continue

            # Standardized Regex matcher for V.O.I.D. filenames
            # Looks for: [Letters/Numbers]_[Numbers]
            # Example: CBA_123456_accounts.xlsx -> Inst: CBA, Batch: 123456
            match = re.match(r'^([A-Za-z0-9]+)_(\d+)', filename)
            
            if match:
                inst_code = match.group(1).upper()
                batch_num = match.group(2)
                
                # Group files under the same batch number
                if batch_num not in current_batches:
                    current_batches[batch_num] = {
                        'institution_code': inst_code,
                        'batch_number': batch_num,
                        'files': []
                    }
                current_batches[batch_num]['files'].append(file_path)

        # Push newly completed batches to the GUI
        for batch_num, batch_data in current_batches.items():
            if batch_num not in self.detected_batches:
                self.detected_batches.add(batch_num)
                
                # Send the batch data to populate the UI panel
                self.message_queue.put({
                    'type': 'batch_detected',
                    'data': batch_data
                })
                
                # Log the detection in the activity feed
                file_count = len(batch_data['files'])
                self.message_queue.put({
                    'type': 'activity',
                    'data': f"Detected: {batch_data['institution_code']} Batch {batch_num} ({file_count} files)"
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
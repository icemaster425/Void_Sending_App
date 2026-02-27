import os
import time
import threading
import re
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from collections import defaultdict

class FileMonitor:
    def __init__(self, folder_path, db_manager, message_queue, file_extensions, batch_length=6):
        self.folder_path = folder_path
        self.db_manager = db_manager
        self.message_queue = message_queue
        self.file_extensions = [ext.lower().strip() for ext in file_extensions]
        self.batch_length = str(batch_length)
        
        self.observer = Observer()
        self.is_running = threading.Event()
        self.lock = threading.Lock()
        
        # Stores files as they arrive: {batch_id: {'excel': path, 'images': [paths], 'inst_code': code}}
        self.detected_batches = defaultdict(lambda: {'excel': None, 'images': [], 'institution_code': None})
        self.batches_in_queue = set()
        self.event_handler = self._create_event_handler()

    def _monitor(self):
        """Internal loop for the observer thread."""
        try:
            self.observer.schedule(self.event_handler, self.folder_path, recursive=False)
            self.observer.start()
            self.message_queue.put({'type': 'activity', 'data': f"Monitoring started: {self.folder_path}"})
            
            while self.is_running.is_set():
                time.sleep(1)
        except Exception as e:
            self.message_queue.put({'type': 'activity', 'data': f"Critical Monitor Error: {str(e)}"})
        finally:
            self.observer.stop()
            self.observer.join()
        
    def start(self):
        """Starts monitoring in a background thread."""
        self.is_running.set()
        self.thread = threading.Thread(target=self._monitor, daemon=True)
        self.thread.start()

    def stop(self):
        """Stops the observer."""
        self.is_running.clear()

    def remove_from_queue(self, batch_number):
        """Cleans up memory once a batch is processed or cancelled."""
        with self.lock:
            if batch_number in self.batches_in_queue:
                self.batches_in_queue.remove(batch_number)
            if batch_number in self.detected_batches:
                del self.detected_batches[batch_number]

    def _create_event_handler(self):
        """Creates the logic for handling file system events."""
        parent = self
        
        class Handler(FileSystemEventHandler):
            def on_created(self, event):
                if event.is_directory:
                    return
                
                filename = os.path.basename(event.src_path)
                ext = os.path.splitext(filename)[1].lower()

                if ext not in parent.file_extensions:
                    return

                batch_number = None
                inst_code = None

                # Pattern for Excel: e.g., 123456_ABC.xlsx 
                excel_pattern = rf'(\d{{{parent.batch_length},}})_([A-Z0-9]+)'
                # Pattern for Images/PDFs: e.g., image_123456.pdf 
                image_pattern = rf'_(\d{{{parent.batch_length},}})'

                if ext in ['.xlsx', '.xls']:
                    match = re.search(excel_pattern, filename, re.IGNORECASE)
                    if match:
                        batch_number, inst_code = match.group(1), match.group(2)
                else:
                    match = re.search(image_pattern, filename)
                    if match:
                        batch_number = match.group(1)
                
                if not batch_number:
                    return
                
                with parent.lock:
                    # Ignore if this batch is already being handled in the UI
                    if batch_number in parent.batches_in_queue:
                        return

                    if ext in ['.xlsx', '.xls']:
                        parent.detected_batches[batch_number]['excel'] = event.src_path
                        parent.detected_batches[batch_number]['institution_code'] = inst_code
                    else:
                        parent.detected_batches[batch_number]['images'].append(event.src_path)

                    data = parent.detected_batches[batch_number]
                    
                    # Logic: A batch is 'ready' when we have both the Excel manifest AND at least one image/PDF
                    if data['excel'] and data['images']:
                        parent.message_queue.put({
                            'type': 'batch_detected',
                            'data': {
                                'batch_number': batch_number,
                                'institution_code': data['institution_code'],
                                'files': [data['excel']] + data['images']
                            }
                        })
                        parent.batches_in_queue.add(batch_number)

        return Handler()
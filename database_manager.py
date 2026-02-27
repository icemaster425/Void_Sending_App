import sqlite3
import os
import sys
from datetime import datetime

class DatabaseManager:
    def __init__(self, db_path=None):
        """
        Initializes the database manager with Multi-User support.
        """
        # If no path is provided, it defaults to the local folder
        if db_path is None:
            if getattr(sys, 'frozen', False):
                script_dir = os.path.dirname(sys.executable)
            else:
                script_dir = os.path.dirname(os.path.abspath(__file__))
            self.db_path = os.path.join(script_dir, 'file_monitor.db')
        else:
            self.db_path = db_path
            
        self.init_database()

    def _get_connection(self):
        """
        Creates a connection with WAL mode enabled for concurrent network access.
        """
        # Increased timeout to 15s to handle potential network latency
        conn = sqlite3.connect(self.db_path, timeout=15)
        # WAL mode allows multiple readers and one writer simultaneously
        conn.execute("PRAGMA journal_mode=WAL;")
        return conn

    def init_database(self):
        """Creates the necessary tables if they do not exist."""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        # Institutions Table 
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS institutions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                county_code TEXT,
                institution_code TEXT NOT NULL UNIQUE,
                email TEXT NOT NULL,
                encryption_key TEXT,
                message TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Email Log Table - Updated for Date/Time tracking 
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS email_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                institution_code TEXT NOT NULL,
                batch_number TEXT NOT NULL,
                recipient_email TEXT,
                subject TEXT,
                attachment_type TEXT,
                attachment_files TEXT,
                sent_date TEXT,
                sent_time TEXT,
                dispatcher_name TEXT,
                record_count INTEGER,
                FOREIGN KEY (institution_code) REFERENCES institutions (institution_code)
            )
        ''')
        conn.commit()
        conn.close()

    def add_sent_email(self, institution_code, batch_number, recipient_email, subject, attachment_type, attachment_files, dispatcher, count=0):
        """Logs a successful draft into history using standardized DD/MM/YYYY format."""
        conn = self._get_connection()
        cursor = conn.cursor()
        # Standardized Date Format: DD/MM/YYYY
        sent_date = datetime.now().strftime('%d/%m/%Y')
        sent_time = datetime.now().strftime('%H:%M:%S')
        files_string = ', '.join(os.path.basename(f) for f in attachment_files)
        
        cursor.execute('''
            INSERT INTO email_log (institution_code, batch_number, recipient_email, subject, 
                                 attachment_type, attachment_files, sent_date, sent_time, 
                                 dispatcher_name, record_count)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (institution_code, batch_number, recipient_email, subject, 
              attachment_type, files_string, sent_date, sent_time, dispatcher, count))
        
        conn.commit()
        conn.close()

    def get_sent_emails(self):
        """Retrieves full history for the UI."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT institution_code, batch_number, sent_date, sent_time, attachment_files, dispatcher_name, record_count
            FROM email_log 
            ORDER BY id DESC
        ''')
        results = cursor.fetchall()
        conn.close()
        return results

    def search_sent_emails(self, search_term):
        """Filters history by Batch, Institution, or Date (DD/MM/YYYY)."""
        conn = self._get_connection()
        cursor = conn.cursor()
        term = f'%{search_term}%'
        cursor.execute('''
            SELECT institution_code, batch_number, sent_date, sent_time, attachment_files, dispatcher_name, record_count
            FROM email_log 
            WHERE batch_number LIKE ? OR institution_code LIKE ? OR sent_date LIKE ?
            ORDER BY id DESC
        ''', (term, term, term))
        results = cursor.fetchall()
        conn.close()
        return results

    # --- Institution Management Methods (Standard CRUD) --- 

    def get_all_institutions(self):
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT county_code, institution_code, email, encryption_key, message FROM institutions ORDER BY institution_code')
        rows = cursor.fetchall()
        conn.close()
        return rows

    def get_institution_by_code(self, institution_code):
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM institutions WHERE institution_code = ?', (institution_code,))
        row = cursor.fetchone()
        conn.close()
        if row:
            return {'county_code': row[1], 'institution_code': row[2], 'email': row[3], 'encryption_key': row[4], 'message': row[5]}
        return None

    def add_institution(self, c_code, i_code, email, key, msg):
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute('INSERT INTO institutions (county_code, institution_code, email, encryption_key, message) VALUES (?, ?, ?, ?, ?)', (c_code, i_code, email, key, msg))
        conn.commit()
        conn.close()

    def update_institution(self, old_code, up_data):
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute('''UPDATE institutions SET county_code=?, institution_code=?, email=?, encryption_key=?, message=?, updated_at=CURRENT_TIMESTAMP 
                          WHERE institution_code=?''', (up_data['county_code'], up_data['institution_code'], up_data['email'], up_data['encryption_key'], up_data['message'], old_code))
        conn.commit()
        conn.close()

    def delete_institution(self, code):
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM institutions WHERE institution_code = ?', (code,))
        conn.commit()
        conn.close()
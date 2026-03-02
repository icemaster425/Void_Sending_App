import sqlite3
import os
from datetime import datetime

class DatabaseManager:
    def __init__(self, db_path):
        """
        Initializes the database manager with a dynamic path provided by 
        the local_config.ini (via the Settings Tab).
        """
        self.db_path = db_path
        self.init_database()

    def _get_connection(self):
        """
        Creates a connection with WAL mode enabled for concurrent network access.
        Timeout of 15 seconds prevents locking crashes if someone else is saving data.
        """
        conn = sqlite3.connect(self.db_path, timeout=15.0)
        conn.execute("PRAGMA journal_mode=WAL;")
        return conn

    def init_database(self):
        """Creates tables and automatically adds missing columns if an older DB is detected."""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        # 1. Base Tables
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
                FOREIGN KEY (institution_code) REFERENCES institutions (institution_code)
            )
        ''')

        # 2. Smart Auto-Migration Logic (Safely adds new columns without crashing)
        cursor.execute("PRAGMA table_info(email_log)")
        columns = [column[1] for column in cursor.fetchall()]
        
        if 'dispatcher_name' not in columns:
            cursor.execute('ALTER TABLE email_log ADD COLUMN dispatcher_name TEXT DEFAULT "Unknown"')
        
        if 'record_count' not in columns:
            cursor.execute('ALTER TABLE email_log ADD COLUMN record_count INTEGER DEFAULT 0')

        conn.commit()
        conn.close()

    def add_sent_email(self, institution_code, batch_number, recipient_email, subject, attachment_type, attachment_files, dispatcher_name="Unknown", record_count=0):
        """Logs a successful dispatch into the shared history."""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        # Enforcing the strictly requested DD/MM/YYYY format
        sent_date = datetime.now().strftime('%d/%m/%Y')
        sent_time = datetime.now().strftime('%H:%M:%S')
        files_string = ', '.join(os.path.basename(f) for f in attachment_files)
        
        cursor.execute('''
            INSERT INTO email_log (institution_code, batch_number, recipient_email, subject, attachment_type, attachment_files, sent_date, sent_time, dispatcher_name, record_count)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (institution_code, batch_number, recipient_email, subject, attachment_type, files_string, sent_date, sent_time, dispatcher_name, record_count))
        
        conn.commit()
        conn.close()

    def get_sent_emails(self):
        conn = self._get_connection()
        cursor = conn.cursor()
        # Sorting by ID descending ensures chronological order even with text-based dates
        cursor.execute('''
            SELECT institution_code, batch_number, sent_date, sent_time, attachment_files, dispatcher_name, record_count 
            FROM email_log 
            ORDER BY id DESC
        ''')
        results = cursor.fetchall()
        conn.close()
        return results

    def get_sent_emails_by_date(self, date_str):
        """Retrieves history for a specific date (DD/MM/YYYY)."""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT institution_code, batch_number, sent_date, sent_time, attachment_files, dispatcher_name, record_count 
            FROM email_log 
            WHERE sent_date = ? 
            ORDER BY id DESC
        ''', (date_str,))
        results = cursor.fetchall()
        conn.close()
        return results

    def search_sent_emails(self, search_term):
        """Searches by Batch Number, Institution Code, or Date."""
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

    # --- Institution Data Methods ---

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

    def add_institution(self, county_code, institution_code, email, encryption_key, message):
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute('INSERT INTO institutions (county_code, institution_code, email, encryption_key, message) VALUES (?, ?, ?, ?, ?)', (county_code, institution_code, email, encryption_key, message))
        conn.commit()
        conn.close()

    def update_institution(self, old_code, updated_data):
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            UPDATE institutions
            SET county_code = ?, institution_code = ?, email = ?, encryption_key = ?, message = ?, updated_at = CURRENT_TIMESTAMP
            WHERE institution_code = ?
        ''', (updated_data['county_code'], updated_data['institution_code'], updated_data['email'], updated_data['encryption_key'], updated_data['message'], old_code))
        conn.commit()
        conn.close()

    def delete_institution(self, institution_code):
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM institutions WHERE institution_code = ?', (institution_code,))
        conn.commit()
        conn.close()
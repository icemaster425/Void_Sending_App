import sqlite3
import os
import time
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
        Creates a connection with WAL mode enabled.
        Timeout increased to 30s for slow network drive resilience.
        """
        conn = sqlite3.connect(self.db_path, timeout=30.0)
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA synchronous=NORMAL;")
        return conn

    def _retry_execute(self, query, params=(), commit=False):
        """High-IQ Network Resilience: Retries on database locks."""
        for attempt in range(5):
            try:
                conn = self._get_connection()
                cursor = conn.cursor()
                cursor.execute(query, params)
                
                if commit:
                    conn.commit()
                    res = cursor.lastrowid
                else:
                    res = cursor.fetchall()
                    
                conn.close()
                return res
            except sqlite3.OperationalError as e:
                if "locked" in str(e).lower() and attempt < 4:
                    time.sleep(1 + attempt) # Incremental backoff
                    continue
                raise e

    def init_database(self):
        """Creates tables and automatically adds missing columns if an older DB is detected."""
        self._retry_execute('''
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
        ''', commit=True)
        
        self._retry_execute('''
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
        ''', commit=True)

        # Smart Auto-Migration Logic
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(email_log)")
        columns = [column[1] for column in cursor.fetchall()]
        
        if 'dispatcher_name' not in columns:
            cursor.execute('ALTER TABLE email_log ADD COLUMN dispatcher_name TEXT DEFAULT "Unknown"')
        
        if 'record_count' not in columns:
            cursor.execute('ALTER TABLE email_log ADD COLUMN record_count INTEGER DEFAULT 0')

        conn.commit()
        conn.close()

    def add_sent_email(self, institution_code, batch_number, recipient_email, subject, attachment_type, attachment_files, dispatcher_name="Unknown", record_count=0):
        sent_date = datetime.now().strftime('%d/%m/%Y')
        sent_time = datetime.now().strftime('%H:%M:%S')
        files_string = ', '.join(os.path.basename(f) for f in attachment_files)
        
        self._retry_execute('''
            INSERT INTO email_log (institution_code, batch_number, recipient_email, subject, attachment_type, attachment_files, sent_date, sent_time, dispatcher_name, record_count)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (institution_code, batch_number, recipient_email, subject, attachment_type, files_string, sent_date, sent_time, dispatcher_name, record_count), commit=True)

    def get_sent_emails(self):
        return self._retry_execute('''
            SELECT institution_code, batch_number, sent_date, sent_time, attachment_files, dispatcher_name, record_count 
            FROM email_log 
            ORDER BY id DESC
        ''')

    def get_sent_emails_by_date(self, date_str):
        return self._retry_execute('''
            SELECT institution_code, batch_number, sent_date, sent_time, attachment_files, dispatcher_name, record_count 
            FROM email_log 
            WHERE sent_date = ? 
            ORDER BY id DESC
        ''', (date_str,))

    def search_sent_emails(self, search_term):
        term = f'%{search_term}%'
        return self._retry_execute('''
            SELECT institution_code, batch_number, sent_date, sent_time, attachment_files, dispatcher_name, record_count 
            FROM email_log 
            WHERE batch_number LIKE ? OR institution_code LIKE ? OR sent_date LIKE ?
            ORDER BY id DESC
        ''', (term, term, term))

    def get_all_institutions(self):
        return self._retry_execute('SELECT county_code, institution_code, email, encryption_key, message FROM institutions ORDER BY institution_code')

    def get_institution_by_code(self, institution_code):
        results = self._retry_execute('SELECT * FROM institutions WHERE institution_code = ?', (institution_code,))
        if results:
            row = results[0]
            return {'county_code': row[1], 'institution_code': row[2], 'email': row[3], 'encryption_key': row[4], 'message': row[5]}
        return None

    def add_institution(self, county_code, institution_code, email, encryption_key, message):
        self._retry_execute('INSERT INTO institutions (county_code, institution_code, email, encryption_key, message) VALUES (?, ?, ?, ?, ?)', 
                            (county_code, institution_code, email, encryption_key, message), commit=True)

    def update_institution(self, old_code, updated_data):
        self._retry_execute('''
            UPDATE institutions
            SET county_code = ?, institution_code = ?, email = ?, encryption_key = ?, message = ?, updated_at = CURRENT_TIMESTAMP
            WHERE institution_code = ?
        ''', (updated_data['county_code'], updated_data['institution_code'], updated_data['email'], updated_data['encryption_key'], updated_data['message'], old_code), commit=True)

    def delete_institution(self, institution_code):
        self._retry_execute('DELETE FROM institutions WHERE institution_code = ?', (institution_code,), commit=True)
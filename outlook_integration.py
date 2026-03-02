import win32com.client
import os
import logging

class OutlookIntegration:
    def __init__(self):
        self.outlook_app = None
        try:
            # Initializes the connection to the local Outlook application
            self.outlook_app = win32com.client.Dispatch("Outlook.Application")
        except Exception as e:
            logging.error(f"Failed to initialize Outlook COM object: {e}")

    def create_draft(self, recipient_email, subject, body, attachment_paths=None):
        """
        Commands Outlook to create and display an email draft.
        Returns: (bool success, str message)
        """
        if not self.outlook_app:
            return False, "Outlook application is not accessible or not installed."

        try:
            # 0 corresponds to an olMailItem (Standard Email)
            mail = self.outlook_app.CreateItem(0) 
            
            mail.To = recipient_email
            mail.Subject = subject
            mail.Body = body

            # Attach the processed and encrypted ZIP files
            if attachment_paths:
                for attachment in attachment_paths:
                    abs_path = os.path.abspath(attachment)
                    if os.path.exists(abs_path):
                        mail.Attachments.Add(abs_path)
                    else:
                        logging.warning(f"Attachment not found: {abs_path}")
                        return False, f"Attachment missing: {os.path.basename(abs_path)}"

            # Display the email on screen so the dispatcher can review it before clicking 'Send'
            # Passing 'False' ensures it doesn't freeze the Python application while open
            mail.Display(False)
            
            # The draft was successfully generated and handed off to Outlook
            return True, "Draft created successfully."
            
        except Exception as e:
            logging.error(f"Error creating Outlook draft: {e}")
            return False, f"Outlook Error: {str(e)}"
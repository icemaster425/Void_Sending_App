import win32com.client
import pythoncom
import os
import logging

class OutlookIntegration:
    def __init__(self):
        self.outlook_app = None

    def create_draft(self, recipient_email, subject, body, attachment_paths=None):
        """
        Commands Outlook to create and display an email draft.
        Returns: (bool success, str message)
        """
        # Critical: Initialize COM for the current thread
        pythoncom.CoInitialize()
        
        try:
            if not self.outlook_app:
                self.outlook_app = win32com.client.Dispatch("Outlook.Application")
            
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

            # Display the email on screen
            mail.Display(False)
            return True, "Draft created successfully."
            
        except Exception as e:
            logging.error(f"Error creating Outlook draft: {e}")
            return False, f"Outlook Error: {str(e)}"
        finally:
            # Clean up COM
            pythoncom.CoUninitialize()
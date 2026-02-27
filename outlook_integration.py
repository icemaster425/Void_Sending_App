import win32com.client as win32
import os

class OutlookIntegration:
    """
    Handles integration with Microsoft Outlook using the MAPI interface.
    
    This class provides functionality to create and display email drafts 
    directly within the user's Outlook instance.
    """
    def __init__(self):
        self.outlook_app = None

    def create_draft(self, to_email, subject, body, attachment_paths=None):
        """
        Creates and displays an Outlook email draft.

        Args:
            to_email (str): The recipient's email address.
            subject (str): The subject line of the email.
            body (str): The main message content.
            attachment_paths (list, optional): Absolute paths to the ZIP files.

        Returns:
            tuple: (bool success, str message)
        """
        try:
            # Connect to the Outlook Application
            self.outlook_app = win32.Dispatch('Outlook.Application')
            
            # Create a new mail item (0 = olMailItem)
            mail_item = self.outlook_app.CreateItem(0)
            
            mail_item.Subject = subject
            mail_item.Body = body
            mail_item.To = to_email
            
            # Attach files if provided
            if attachment_paths:
                for path in attachment_paths:
                    if os.path.exists(path):
                        mail_item.Attachments.Add(path)
                    else:
                        return False, f"Attachment not found: {os.path.basename(path)}"
            
            # Display the draft instead of sending it immediately
            # This allows the user to perform a final manual check.
            mail_item.Display() 
            
            return True, "Outlook draft created successfully."
            
        except Exception as e:
            # Usually fails if Outlook is not installed or blocked by security
            return False, f"Outlook Error: {str(e)}"
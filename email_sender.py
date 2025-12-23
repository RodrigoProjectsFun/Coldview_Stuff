"""
Outlook Email Sender Module
============================
Sends emails with attachments via Microsoft Outlook using COM automation.

Requirements:
- Windows OS
- Microsoft Outlook desktop app installed and configured
- pywin32 package (pip install pywin32)
"""

import os
import json
from datetime import datetime

try:
    import win32com.client
except ImportError:
    print("ERROR: pywin32 not installed. Run: pip install pywin32")
    raise


def load_email_config(config_path: str = "config.json") -> dict:
    """Load email configuration from config file."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    full_path = os.path.join(script_dir, config_path)
    
    with open(full_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    return config.get("email", {})


def send_email_with_attachment(
    attachment_path: str,
    to: str = None,
    cc: str = None,
    subject: str = None,
    body: str = None,
    config_path: str = "config.json"
) -> bool:
    """
    Send an email via Outlook with an attachment.
    
    Args:
        attachment_path: Full path to the file to attach.
        to: Recipient email (overrides config if provided).
        cc: CC recipients (overrides config if provided).
        subject: Email subject (overrides config if provided).
        body: Email body (overrides config if provided).
        config_path: Path to config file with email settings.
    
    Returns:
        bool: True if email sent successfully, False otherwise.
    """
    # Load config
    email_config = load_email_config(config_path)
    
    # Get filename for template substitution
    filename = os.path.basename(attachment_path)
    current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Use provided values or fallback to config
    recipient = to or email_config.get("to", "")
    cc_list = cc or email_config.get("cc", "")
    
    # Process subject template
    subject_template = subject or email_config.get("subject_template", "New File: {filename}")
    email_subject = subject_template.format(filename=filename, date=current_date)
    
    # Process body template
    body_template = body or email_config.get("body_template", "File attached: {filename}")
    email_body = body_template.format(filename=filename, date=current_date)
    
    if not recipient:
        print("ERROR: No recipient email specified in config or method call.")
        return False
    
    if not os.path.exists(attachment_path):
        print(f"ERROR: Attachment file not found: {attachment_path}")
        return False
    
    try:
        print(f"Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = MailItem
        
        mail.To = recipient
        if cc_list:
            mail.CC = cc_list
        mail.Subject = email_subject
        mail.Body = email_body
        
        # Attach the file
        print(f"Attaching file: {filename}")
        mail.Attachments.Add(attachment_path)
        
        # Send the email
        mail.Send()
        
        print(f"✓ Email sent successfully to: {recipient}")
        if cc_list:
            print(f"  CC: {cc_list}")
        print(f"  Subject: {email_subject}")
        print(f"  Attachment: {filename}")
        
        return True
        
    except Exception as e:
        print(f"ERROR: Failed to send email: {e}")
        return False


def test_outlook_connection() -> bool:
    """Test if Outlook is accessible via COM."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        print("✓ Outlook connection successful")
        return True
    except Exception as e:
        print(f"✗ Outlook connection failed: {e}")
        return False


if __name__ == "__main__":
    # Quick test
    print("Testing Outlook connection...")
    test_outlook_connection()

#!/usr/bin/env python3
"""
Gmail Manager for CoastRev VM Migration
Replaces win32com.client Outlook operations with Gmail API
"""
import os
import re
import base64
import requests
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Import our utilities
import sys
sys.path.append('../utils')
sys.path.append('../config')
from google_drive_manager import drive_manager
from paths import path_manager

class GmailManager:
    """Manages email operations with Gmail API"""
    
    def __init__(self, credentials_path=None):
        self.credentials_path = credentials_path or "/home/user/coastrev/creds/blissful-mantis-471621-p1-ed8f8bd5470b.json"
        self.service = self._setup_service()
        
    def _setup_service(self):
        """Initialize Gmail service with service account"""
        try:
            credentials = Credentials.from_service_account_file(
                self.credentials_path,
                scopes=['https://www.googleapis.com/auth/gmail.readonly']
            )
            return build('gmail', 'v1', credentials=credentials)
        except Exception as e:
            print(f"Failed to initialize Gmail service: {e}")
            return None
    
    def search_messages(self, query, max_results=50):
        """Search for messages matching query"""
        try:
            if not self.service:
                print("❌ Gmail service not initialized")
                return []
                
            results = self.service.users().messages().list(
                userId='me', 
                q=query, 
                maxResults=max_results
            ).execute()
            
            messages = results.get('messages', [])
            return messages
            
        except Exception as e:
            print(f"❌ Failed to search messages: {e}")
            return []
    
    def get_message_details(self, message_id):
        """Get full message details"""
        try:
            message = self.service.users().messages().get(
                userId='me', 
                id=message_id,
                format='full'
            ).execute()
            
            return message
            
        except Exception as e:
            print(f"❌ Failed to get message {message_id}: {e}")
            return None
    
    def extract_attachments(self, message, save_to_folder_id=None):
        """Extract attachments from a message and upload to Google Drive"""
        try:
            attachments = []
            
            if 'payload' in message:
                payload = message['payload']
                
                if 'parts' in payload:
                    for part in payload['parts']:
                        if part.get('filename'):
                            attachment_data = self._get_attachment_data(message['id'], part['body']['attachmentId'])
                            if attachment_data:
                                attachment = {
                                    'filename': part['filename'],
                                    'data': attachment_data,
                                    'mime_type': part['mimeType']
                                }
                                attachments.append(attachment)
                
                # Handle single attachment (no parts structure)
                elif payload.get('filename'):
                    attachment_data = self._get_attachment_data(message['id'], payload['body']['attachmentId'])
                    if attachment_data:
                        attachment = {
                            'filename': payload['filename'],
                            'data': attachment_data,
                            'mime_type': payload['mimeType']
                        }
                        attachments.append(attachment)
            
            # Save attachments to Google Drive if folder specified
            saved_files = []
            if save_to_folder_id and attachments:
                for attachment in attachments:
                    file_id = self._save_attachment_to_drive(attachment, save_to_folder_id)
                    if file_id:
                        saved_files.append({
                            'filename': attachment['filename'],
                            'drive_file_id': file_id
                        })
                        
            return saved_files
            
        except Exception as e:
            print(f"❌ Failed to extract attachments: {e}")
            return []
    
    def _get_attachment_data(self, message_id, attachment_id):
        """Get attachment data"""
        try:
            attachment = self.service.users().messages().attachments().get(
                userId='me',
                messageId=message_id,
                id=attachment_id
            ).execute()
            
            data = base64.urlsafe_b64decode(attachment['data'])
            return data
            
        except Exception as e:
            print(f"❌ Failed to get attachment data: {e}")
            return None
    
    def _save_attachment_to_drive(self, attachment, folder_id):
        """Save attachment directly to Google Drive"""
        try:
            import tempfile
            
            # Create temporary file
            with tempfile.NamedTemporaryFile(delete=False, suffix=f"_{attachment['filename']}") as temp_file:
                temp_file.write(attachment['data'])
                temp_path = temp_file.name
                
            # Upload to Drive
            file_id = drive_manager.upload_file(temp_path, folder_id, attachment['filename'])
            
            # Clean up temp file
            os.unlink(temp_path)
            
            if file_id:
                print(f"✅ Saved attachment to Drive: {attachment['filename']}")
            
            return file_id
            
        except Exception as e:
            print(f"❌ Failed to save attachment to Drive: {e}")
            return None
    
    def process_outlook_attachments(self, property_code, subject_keyword, date_filter=None):
        """
        Replace win32com.client Outlook attachment processing
        Downloads attachments from emails matching subject keyword
        """
        try:
            if date_filter is None:
                date_filter = datetime.now().strftime("%Y/%m/%d")
                
            # Build search query
            query = f'subject:"{subject_keyword}" after:{date_filter}'
            print(f"🔍 Searching for emails: {query}")
            
            # Search messages
            messages = self.search_messages(query)
            
            if not messages:
                print(f"⚠️ No emails found with subject containing '{subject_keyword}' from {date_filter}")
                return False
                
            # Get property's email attachments folder
            property_folder_id = path_manager.get_property_drive_folder(property_code)
            if not property_folder_id:
                print(f"❌ No Drive folder configured for property {property_code}")
                return False
            
            total_attachments = 0
            
            # Process each message
            for message in messages[:5]:  # Limit to latest 5 messages
                message_details = self.get_message_details(message['id'])
                if message_details:
                    saved_files = self.extract_attachments(message_details, property_folder_id)
                    total_attachments += len(saved_files)
                    
                    for file_info in saved_files:
                        print(f"✅ Processed: {file_info['filename']} -> Drive ID: {file_info['drive_file_id']}")
            
            if total_attachments > 0:
                print(f"✅ Successfully processed {total_attachments} attachments for {property_code}")
                return True
            else:
                print(f"⚠️ No attachments found in matching emails for {property_code}")
                return False
                
        except Exception as e:
            print(f"❌ Error processing Outlook attachments for {property_code}: {e}")
            return False
    
    def process_cloudbeds_emails(self, property_code, subject_filter="rooms sold and occupancy"):
        """
        Replace Cloudbeds email link extraction and CSV download
        Processes Cloudbeds notification emails with download links
        """
        try:
            today = datetime.now().strftime("%Y/%m/%d")
            
            # Search for Cloudbeds emails
            query = f'from:noreply@cloudbeds.com subject:"{subject_filter}" after:{today}'
            print(f"🔍 Searching for Cloudbeds emails: {query}")
            
            messages = self.search_messages(query)
            
            if not messages:
                print(f"⚠️ No Cloudbeds emails found with subject '{subject_filter}' from today")
                return False
                
            # Get property's email attachments folder
            property_folder_id = path_manager.get_property_drive_folder(property_code)
            if not property_folder_id:
                print(f"❌ No Drive folder configured for property {property_code}")
                return False
            
            seen_links = set()
            downloaded_files = 0
            
            # Process each message
            for message in messages:
                message_details = self.get_message_details(message['id'])
                if not message_details:
                    continue
                    
                # Extract HTML body
                html_body = self._extract_html_body(message_details)
                if not html_body:
                    continue
                
                # Extract download links from HTML
                download_links = self._extract_cloudbeds_links(html_body)
                
                for link_info in download_links:
                    if link_info['url'] in seen_links:
                        continue
                        
                    seen_links.add(link_info['url'])
                    
                    # Download and save to Drive
                    if self._download_and_save_csv(link_info, property_code, property_folder_id):
                        downloaded_files += 1
            
            if downloaded_files > 0:
                print(f"✅ Successfully downloaded {downloaded_files} Cloudbeds reports for {property_code}")
                return True
            else:
                print(f"⚠️ No valid Cloudbeds download links found for {property_code}")
                return False
                
        except Exception as e:
            print(f"❌ Error processing Cloudbeds emails for {property_code}: {e}")
            return False
    
    def _extract_html_body(self, message):
        """Extract HTML body from message"""
        try:
            payload = message.get('payload', {})
            
            # Check if message has parts
            if 'parts' in payload:
                for part in payload['parts']:
                    if part.get('mimeType') == 'text/html':
                        data = part.get('body', {}).get('data', '')
                        if data:
                            return base64.urlsafe_b64decode(data).decode('utf-8')
            
            # Check main body
            elif payload.get('mimeType') == 'text/html':
                data = payload.get('body', {}).get('data', '')
                if data:
                    return base64.urlsafe_b64decode(data).decode('utf-8')
            
            return None
            
        except Exception as e:
            print(f"❌ Failed to extract HTML body: {e}")
            return None
    
    def _extract_cloudbeds_links(self, html_body):
        """Extract Cloudbeds download links from HTML"""
        try:
            links = []
            
            # Extract report title
            title_match = re.search(r'<a[^>]*>([^<]+)</a>', html_body, re.IGNORECASE)
            report_title = title_match.group(1).strip() if title_match else "cloudbeds_report"
            
            # Extract download link
            link_matches = re.findall(r'href="(https://link\.cloudbeds\.com/[^"]+)"', html_body)
            
            for link in link_matches:
                links.append({
                    'url': link,
                    'title': report_title
                })
                
            return links
            
        except Exception as e:
            print(f"❌ Failed to extract Cloudbeds links: {e}")
            return []
    
    def _download_and_save_csv(self, link_info, property_code, folder_id):
        """Download CSV from Cloudbeds and save to Drive"""
        try:
            # Download the CSV
            response = requests.get(link_info['url'], timeout=30)
            if response.status_code != 200:
                print(f"❌ Failed to download from {link_info['url']}")
                return False
            
            # Sanitize filename
            clean_title = re.sub(r'[\\/*?:"<>|]', "", link_info['title']).strip().replace(" ", "_")
            today_str = datetime.now().strftime("%Y-%m-%d")
            filename = f"{clean_title}_{property_code}_{today_str}.csv"
            
            # Save to temporary file then upload to Drive
            import tempfile
            with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as temp_file:
                temp_file.write(response.content)
                temp_path = temp_file.name
            
            # Upload to Drive
            file_id = drive_manager.upload_file(temp_path, folder_id, filename)
            
            # Clean up
            os.unlink(temp_path)
            
            if file_id:
                print(f"✅ Downloaded and saved: {filename} -> Drive ID: {file_id}")
                return True
            else:
                print(f"❌ Failed to upload {filename} to Drive")
                return False
                
        except Exception as e:
            print(f"❌ Error downloading/saving CSV: {e}")
            return False

# Global instance
gmail_manager = GmailManager()

# Convenience functions
def download_outlook_attachments(property_code, subject_keyword):
    """Download Outlook email attachments for property"""
    return gmail_manager.process_outlook_attachments(property_code, subject_keyword)

def download_cloudbeds_reports(property_code):
    """Download Cloudbeds reports for property"""
    return gmail_manager.process_cloudbeds_emails(property_code)

# Test the Gmail manager
if __name__ == "__main__":
    print("Gmail Manager Test")
    print("="*40)
    
    if gmail_manager.service:
        print("✅ Gmail service initialized successfully")
        
        # Test search
        messages = gmail_manager.search_messages("subject:test", max_results=1)
        print(f"✅ Found {len(messages)} test messages")
        
    else:
        print("❌ Failed to initialize Gmail service")
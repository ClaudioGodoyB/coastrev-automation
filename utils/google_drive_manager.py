#!/usr/bin/env python3
"""
Google Drive File Manager for CoastRev VM Migration
Replaces local file operations with Google Drive API calls
"""
import os
import io
import json
from datetime import datetime
from pathlib import Path
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

# Import path configuration
import sys
sys.path.append('../config')
from paths import path_manager

class GoogleDriveManager:
    """Manages file operations with Google Drive API"""
    
    def __init__(self, credentials_path=None):
        self.credentials_path = credentials_path or "/home/user/coastrev/creds/blissful-mantis-471621-p1-ed8f8bd5470b.json"
        self.service = self._setup_service()
        
    def _setup_service(self):
        """Initialize Google Drive service"""
        try:
            credentials = Credentials.from_service_account_file(
                self.credentials_path,
                scopes=['https://www.googleapis.com/auth/drive']
            )
            return build('drive', 'v3', credentials=credentials)
        except Exception as e:
            print(f"Failed to initialize Google Drive service: {e}")
            return None
    
    def create_folder(self, name, parent_id=None):
        """Create a folder in Google Drive"""
        folder_metadata = {
            'name': name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        
        if parent_id:
            folder_metadata['parents'] = [parent_id]
            
        try:
            folder = self.service.files().create(body=folder_metadata).execute()
            print(f"✅ Created folder: {name} (ID: {folder['id']})")
            return folder['id']
        except Exception as e:
            print(f"❌ Failed to create folder {name}: {e}")
            return None
    
    def upload_file(self, local_file_path, parent_folder_id=None, drive_file_name=None):
        """Upload a file to Google Drive"""
        try:
            if not os.path.exists(local_file_path):
                print(f"Local file not found: {local_file_path}")
                return None
                
            file_name = drive_file_name or os.path.basename(local_file_path)
            
            file_metadata = {'name': file_name}
            if parent_folder_id:
                file_metadata['parents'] = [parent_folder_id]
                
            media = MediaFileUpload(local_file_path, resumable=True)
            
            file = self.service.files().create(
                body=file_metadata,
                media_body=media
            ).execute()
            
            print(f"✅ Uploaded: {file_name} (ID: {file['id']})")
            return file['id']
            
        except Exception as e:
            print(f"❌ Failed to upload {local_file_path}: {e}")
            return None
    
    def download_file(self, file_id, local_file_path):
        """Download a file from Google Drive"""
        try:
            # Get file metadata first
            file_metadata = self.service.files().get(fileId=file_id).execute()
            
            # Create local directory if needed
            os.makedirs(os.path.dirname(local_file_path), exist_ok=True)
            
            # Download file
            request = self.service.files().get_media(fileId=file_id)
            
            with io.FileIO(local_file_path, 'wb') as fh:
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while done is False:
                    status, done = downloader.next_chunk()
                    
            print(f"✅ Downloaded: {file_metadata['name']} -> {local_file_path}")
            return True
            
        except Exception as e:
            print(f"❌ Failed to download file {file_id}: {e}")
            return False
    
    def list_files_in_folder(self, folder_id, file_pattern=None):
        """List files in a Google Drive folder"""
        try:
            query = f"'{folder_id}' in parents and trashed=false"
            if file_pattern:
                query += f" and name contains '{file_pattern}'"
                
            results = self.service.files().list(q=query, fields="files(id,name,modifiedTime)").execute()
            files = results.get('files', [])
            
            return files
            
        except Exception as e:
            print(f"❌ Failed to list files in folder {folder_id}: {e}")
            return []
    
    def find_latest_file(self, folder_id, pattern):
        """Find the latest file matching pattern in folder"""
        files = self.list_files_in_folder(folder_id, pattern)
        
        if not files:
            return None
            
        # Sort by modification time (newest first)
        files.sort(key=lambda x: x['modifiedTime'], reverse=True)
        return files[0]
    
    def get_or_create_dated_folder(self, parent_folder_id, folder_prefix, date_str=None):
        """Get or create a dated folder (e.g., Extract 2025-09-10)"""
        if date_str is None:
            date_str = datetime.today().strftime("%Y-%m-%d")
            
        folder_name = f"{folder_prefix} {date_str}"
        
        # Check if folder already exists
        existing_files = self.list_files_in_folder(parent_folder_id)
        for file in existing_files:
            if file['name'] == folder_name:
                print(f"✅ Found existing folder: {folder_name}")
                return file['id']
        
        # Create new folder if not found
        return self.create_folder(folder_name, parent_folder_id)
    
    def process_csv_from_downloads(self, property_code, pattern):
        """
        Process CSV files from downloads staging area for a specific property
        Replaces the Windows Downloads folder logic
        """
        try:
            # Get downloads staging folder
            staging_folder_id = path_manager.get_drive_folder_id("downloads_staging")
            if not staging_folder_id:
                print("❌ Downloads staging folder not configured")
                return False
                
            # Find matching CSV files
            csv_files = self.list_files_in_folder(staging_folder_id, pattern)
            csv_files = [f for f in csv_files if f['name'].lower().endswith('.csv')]
            
            if not csv_files:
                print(f"No CSV files found matching pattern: {pattern}")
                return False
                
            # Get today's extract folder for the property
            extracts_folder_id = path_manager.get_drive_folder_id("data_inputs") 
            date_str = datetime.today().strftime("%Y-%m-%d")
            dated_folder_id = self.get_or_create_dated_folder(extracts_folder_id, "Extract", date_str)
            
            if not dated_folder_id:
                print("❌ Failed to create/find dated extract folder")
                return False
                
            # Create property folder within dated folder
            property_folder_id = self.get_or_create_property_folder(dated_folder_id, property_code)
            
            # Copy files to property folder
            copied_count = 0
            for csv_file in csv_files:
                # Check if file is from today (simplified - you might want to add date filtering)
                if self.copy_file_to_folder(csv_file['id'], property_folder_id):
                    copied_count += 1
                    
            print(f"✅ Processed {copied_count} CSV files for {property_code}")
            return copied_count > 0
            
        except Exception as e:
            print(f"❌ Error processing CSV files for {property_code}: {e}")
            return False
    
    def copy_file_to_folder(self, source_file_id, destination_folder_id):
        """Copy a file to another folder in Google Drive"""
        try:
            # Get source file metadata
            source_file = self.service.files().get(fileId=source_file_id).execute()
            
            # Create copy
            copy_metadata = {
                'name': source_file['name'],
                'parents': [destination_folder_id]
            }
            
            copied_file = self.service.files().copy(
                fileId=source_file_id,
                body=copy_metadata
            ).execute()
            
            print(f"✅ Copied: {source_file['name']} -> folder {destination_folder_id}")
            return copied_file['id']
            
        except Exception as e:
            print(f"❌ Failed to copy file {source_file_id}: {e}")
            return None
    
    def get_or_create_property_folder(self, parent_folder_id, property_code):
        """Get or create property folder within parent folder"""
        # Check if property folder already exists
        existing_files = self.list_files_in_folder(parent_folder_id)
        for file in existing_files:
            if file['name'] == property_code:
                return file['id']
                
        # Create new property folder
        return self.create_folder(property_code, parent_folder_id)
    
    def sync_local_to_drive(self, local_dir, drive_folder_id):
        """Sync local directory to Google Drive folder"""
        try:
            if not os.path.exists(local_dir):
                print(f"Local directory not found: {local_dir}")
                return False
                
            uploaded_count = 0
            for root, dirs, files in os.walk(local_dir):
                for file in files:
                    local_file_path = os.path.join(root, file)
                    
                    # Skip backup files and hidden files
                    if file.startswith('.') or 'backup' in file.lower():
                        continue
                        
                    if self.upload_file(local_file_path, drive_folder_id):
                        uploaded_count += 1
                        
            print(f"✅ Synced {uploaded_count} files to Google Drive")
            return True
            
        except Exception as e:
            print(f"❌ Error syncing {local_dir} to Drive: {e}")
            return False

# Global instance for easy importing
drive_manager = GoogleDriveManager()

# Convenience functions
def upload_file(local_path, folder_id=None, drive_name=None):
    """Quick file upload"""
    return drive_manager.upload_file(local_path, folder_id, drive_name)

def download_file(file_id, local_path):
    """Quick file download"""
    return drive_manager.download_file(file_id, local_path)

def get_property_folder(property_code):
    """Get email attachments folder for property"""
    return path_manager.get_property_drive_folder(property_code)

def create_dated_folder(parent_id, prefix, date_str=None):
    """Create dated folder"""
    return drive_manager.get_or_create_dated_folder(parent_id, prefix, date_str)

# Test the Google Drive manager
if __name__ == "__main__":
    print("Google Drive Manager Test")
    print("="*40)
    
    if drive_manager.service:
        print("✅ Google Drive service initialized successfully")
        
        # Test folder listing
        root_folder_id = path_manager.get_drive_folder_id("root")
        files = drive_manager.list_files_in_folder(root_folder_id)
        print(f"✅ Found {len(files)} items in root folder")
        
        # Test property folder lookup
        bwof_folder = get_property_folder("BWOF")
        print(f"✅ BWOF folder ID: {bwof_folder}")
        
    else:
        print("❌ Failed to initialize Google Drive service")
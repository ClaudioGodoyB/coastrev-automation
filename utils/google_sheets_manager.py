#!/usr/bin/env python3
"""
Google Sheets Manager for CoastRev VM Migration
Replaces xlwings and win32com.client Excel operations with Google Sheets API
"""
import os
import re
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Import our utilities
import sys
sys.path.append('../config')
from paths import path_manager

class GoogleSheetsManager:
    """Manages spreadsheet operations with Google Sheets API"""
    
    def __init__(self, credentials_path=None):
        self.credentials_path = credentials_path or "/home/user/coastrev/creds/blissful-mantis-471621-p1-ed8f8bd5470b.json"
        self.service = self._setup_service()
        
    def _setup_service(self):
        """Initialize Google Sheets service"""
        try:
            credentials = Credentials.from_service_account_file(
                self.credentials_path,
                scopes=['https://www.googleapis.com/auth/spreadsheets']
            )
            return build('sheets', 'v4', credentials=credentials)
        except Exception as e:
            print(f"Failed to initialize Google Sheets service: {e}")
            return None
    
    def get_cell_value(self, spreadsheet_id, cell_address):
        """
        Get value from a specific cell
        cell_address: e.g., 'B9', 'A1', etc.
        """
        try:
            range_name = f"Sheet1!{cell_address}"
            
            result = self.service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            if values and values[0]:
                return values[0][0]
            else:
                return None
                
        except Exception as e:
            print(f"❌ Failed to get cell {cell_address}: {e}")
            return None
    
    def get_range_values(self, spreadsheet_id, range_address):
        """
        Get values from a range of cells
        range_address: e.g., 'A1:Z50', 'B9:F15', etc.
        """
        try:
            range_name = f"Sheet1!{range_address}"
            
            result = self.service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            return values
            
        except Exception as e:
            print(f"❌ Failed to get range {range_address}: {e}")
            return []
    
    def update_cell_value(self, spreadsheet_id, cell_address, value):
        """Update a single cell with a value"""
        try:
            range_name = f"Sheet1!{cell_address}"
            
            body = {
                'values': [[value]]
            }
            
            result = self.service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption='RAW',
                body=body
            ).execute()
            
            print(f"✅ Updated cell {cell_address} with value: {value}")
            return True
            
        except Exception as e:
            print(f"❌ Failed to update cell {cell_address}: {e}")
            return False
    
    def update_range_values(self, spreadsheet_id, range_address, values):
        """Update a range of cells with values"""
        try:
            range_name = f"Sheet1!{range_address}"
            
            body = {
                'values': values
            }
            
            result = self.service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption='RAW',
                body=body
            ).execute()
            
            print(f"✅ Updated range {range_address}")
            return True
            
        except Exception as e:
            print(f"❌ Failed to update range {range_address}: {e}")
            return False
    
    def create_spreadsheet(self, title, folder_id=None):
        """Create a new Google Spreadsheet"""
        try:
            spreadsheet = {
                'properties': {
                    'title': title
                }
            }
            
            spreadsheet = self.service.spreadsheets().create(
                body=spreadsheet
            ).execute()
            
            spreadsheet_id = spreadsheet['spreadsheetId']
            
            # Move to folder if specified
            if folder_id:
                from google_drive_manager import drive_manager
                drive_manager.service.files().update(
                    fileId=spreadsheet_id,
                    addParents=folder_id,
                    removeParents='root'
                ).execute()
                
            print(f"✅ Created spreadsheet: {title} (ID: {spreadsheet_id})")
            return spreadsheet_id
            
        except Exception as e:
            print(f"❌ Failed to create spreadsheet {title}: {e}")
            return None
    
    def generate_html_from_sheet(self, spreadsheet_id, template_path, output_path, cell_mappings):
        """
        Generate HTML from Google Sheet data using template and cell mappings
        Replaces the xlwings HTML generation functionality
        """
        try:
            # Load HTML template
            if not os.path.exists(template_path):
                print(f"❌ Template not found: {template_path}")
                return False
                
            with open(template_path, 'r', encoding='utf-8') as f:
                html_template = f.read()
            
            # Build placeholder map from cell mappings
            placeholder_map = {}
            
            for placeholder, cell_address in cell_mappings.items():
                cell_value = self.get_cell_value(spreadsheet_id, cell_address)
                
                # Handle different value types based on placeholder name
                if cell_value is not None:
                    formatted_value = self._format_cell_value(placeholder, cell_value)
                    placeholder_map[placeholder] = formatted_value
                else:
                    # Default values for missing data
                    placeholder_map[placeholder] = self._get_default_value(placeholder)
            
            # Replace placeholders in template
            html_content = html_template
            for placeholder, value in placeholder_map.items():
                html_content = html_content.replace(placeholder, str(value))
            
            # Ensure output directory exists
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Write HTML file
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            print(f"✅ Generated HTML report: {output_path}")
            return True
            
        except Exception as e:
            print(f"❌ Failed to generate HTML from sheet: {e}")
            return False
    
    def _format_cell_value(self, placeholder, value):
        """Format cell value based on placeholder pattern"""
        try:
            # Convert to appropriate type
            if value == '':
                return self._get_default_value(placeholder)
            
            # Percentage formatting
            if '%' in placeholder or 'percent' in placeholder.lower():
                try:
                    num_value = float(value)
                    return f"{round(num_value * 100)}%"
                except:
                    return "0%"
            
            # Dollar formatting  
            if '$' in placeholder or 'dollar' in placeholder.lower() or 'revenue' in placeholder.lower():
                try:
                    num_value = float(value)
                    return f"${int(num_value)}"
                except:
                    return "$0"
            
            # Integer formatting
            if any(keyword in placeholder.lower() for keyword in ['count', 'number', 'rooms', 'nights']):
                try:
                    return str(int(float(value)))
                except:
                    return "0"
            
            # Default: return as string
            return str(value)
            
        except Exception as e:
            print(f"Warning: Failed to format value for {placeholder}: {e}")
            return str(value)
    
    def _get_default_value(self, placeholder):
        """Get appropriate default value based on placeholder type"""
        if '%' in placeholder or 'percent' in placeholder.lower():
            return "0%"
        elif '$' in placeholder or 'dollar' in placeholder.lower() or 'revenue' in placeholder.lower():
            return "$0"
        elif any(keyword in placeholder.lower() for keyword in ['count', 'number', 'rooms', 'nights']):
            return "0"
        else:
            return ""
    
    def copy_excel_to_sheets(self, excel_file_path, folder_id=None):
        """
        Copy an Excel file to Google Sheets
        Returns the new spreadsheet ID
        """
        try:
            from google_drive_manager import drive_manager
            
            # Upload Excel file to Drive first
            file_id = drive_manager.upload_file(excel_file_path, folder_id)
            if not file_id:
                return None
            
            # Convert to Google Sheets
            copied_file = drive_manager.service.files().copy(
                fileId=file_id,
                body={
                    'name': os.path.splitext(os.path.basename(excel_file_path))[0],
                    'mimeType': 'application/vnd.google-apps.spreadsheet'
                }
            ).execute()
            
            # Delete the original Excel file from Drive
            drive_manager.service.files().delete(fileId=file_id).execute()
            
            spreadsheet_id = copied_file['id']
            print(f"✅ Converted Excel to Google Sheets: {spreadsheet_id}")
            return spreadsheet_id
            
        except Exception as e:
            print(f"❌ Failed to copy Excel to Sheets: {e}")
            return None
    
    def get_spreadsheet_info(self, spreadsheet_id):
        """Get basic information about a spreadsheet"""
        try:
            spreadsheet = self.service.spreadsheets().get(
                spreadsheetId=spreadsheet_id
            ).execute()
            
            return {
                'title': spreadsheet['properties']['title'],
                'sheets': [sheet['properties']['title'] for sheet in spreadsheet['sheets']]
            }
            
        except Exception as e:
            print(f"❌ Failed to get spreadsheet info: {e}")
            return None

# Global instance
sheets_manager = GoogleSheetsManager()

# Convenience functions
def get_cell_value(spreadsheet_id, cell_address):
    """Quick cell value lookup"""
    return sheets_manager.get_cell_value(spreadsheet_id, cell_address)

def update_cell_value(spreadsheet_id, cell_address, value):
    """Quick cell value update"""
    return sheets_manager.update_cell_value(spreadsheet_id, cell_address, value)

def generate_html_report(spreadsheet_id, template_path, output_path, cell_mappings):
    """Quick HTML report generation"""
    return sheets_manager.generate_html_from_sheet(spreadsheet_id, template_path, output_path, cell_mappings)

# Test the Google Sheets manager
if __name__ == "__main__":
    print("Google Sheets Manager Test")
    print("="*40)
    
    if sheets_manager.service:
        print("✅ Google Sheets service initialized successfully")
        
        # Test basic functionality - create a test spreadsheet
        test_sheet_id = sheets_manager.create_spreadsheet("CoastRev Test Sheet")
        if test_sheet_id:
            print(f"✅ Test spreadsheet created: {test_sheet_id}")
            
            # Test cell operations
            sheets_manager.update_cell_value(test_sheet_id, "A1", "Test Value")
            value = sheets_manager.get_cell_value(test_sheet_id, "A1") 
            print(f"✅ Cell test result: {value}")
        
    else:
        print("❌ Failed to initialize Google Sheets service")
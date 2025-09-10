#!/usr/bin/env python3
"""
BWOF HTML Generation - VM Compatible Version
Replaces xlwings with Google Sheets API
Property: BWOF
"""
import os
import sys
from datetime import datetime

# Add utils to path for imports
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'utils'))
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'config'))

from google_sheets_manager import sheets_manager
from google_drive_manager import drive_manager
from paths import path_manager

# BWOF CONFIGURATION
PROPERTY_CODE = "BWOF"
# Note: This should be the Google Sheets ID of the BWOF template
# Replace with actual spreadsheet ID after template is uploaded to Google Sheets
SPREADSHEET_ID = "YOUR_BWOF_SPREADSHEET_ID_HERE"  # Update this!

def generate_bwof_html():
    """Generate BWOF HTML report from Google Sheets data"""
    try:
        print(f"🏨 Generating HTML report for {PROPERTY_CODE}")
        print("="*50)
        
        # Initialize managers
        if not sheets_manager.service:
            print("❌ Google Sheets service not available")
            return False
            
        if not drive_manager.service:
            print("❌ Google Drive service not available") 
            return False
        
        # Define paths
        current_date = datetime.now().strftime("%Y-%m-%d")
        template_path = path_manager.convert_windows_path(r'C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Templates\Daily_Pickup_Summary_Template.html')
        output_folder = path_manager.convert_windows_path(rf'C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Details\Daily Detail {current_date}\Misc')
        output_file = os.path.join(output_folder, f"Daily_Pickup_Summary_{PROPERTY_CODE}.html")
        
        # Define cell mappings (maps HTML placeholders to Google Sheets cell addresses)
        # This maps the original xlwings cell references to Google Sheets A1 notation
        cell_mappings = {
            # Row 10 (original row 9 in xlwings)
            "{{value_r10_c2}}": "B9",
            "{{value_r10_c3}}": "C9", 
            "{{value_r10_c4}}": "D9",
            "{{value_r10_c5}}": "E9",
            "{{value_r10_c6}}": "F9",
            "{{value_r10_c8}}": "H9",
            "{{value_r10_c9}}": "I9",
            "{{value_r10_c10}}": "J9",
            "{{value_r10_c11}}": "K9",
            "{{value_r10_c12}}": "L9",
            "{{value_r10_c14}}": "N9",
            "{{value_r10_c15}}": "O9",
            "{{value_r10_c16}}": "P9",
            "{{value_r10_c17}}": "Q9",
            "{{value_r10_c18}}": "R9",
            "{{value_r10_c19}}": "S9",
            "{{value_r10_c21}}": "U9",
            "{{value_r10_c22}}": "V9",
            "{{value_r10_c23}}": "W9",
            "{{value_r10_c24}}": "X9",
            "{{value_r10_c25}}": "Y9",
            "{{value_r10_c26}}": "Z9",
            
            # Row 11 (original row 10 in xlwings)
            "{{value_r11_c2}}": "B10",
            "{{value_r11_c3}}": "C10",
            "{{value_r11_c4}}": "D10", 
            "{{value_r11_c5}}": "E10",
            "{{value_r11_c6}}": "F10",
            "{{value_r11_c8}}": "H10",
            "{{value_r11_c9}}": "I10",
            "{{value_r11_c10}}": "J10",
            "{{value_r11_c11}}": "K10",
            "{{value_r11_c12}}": "L10",
            "{{value_r11_c14}}": "N10",
            "{{value_r11_c15}}": "O10",
            "{{value_r11_c16}}": "P10",
            "{{value_r11_c17}}": "Q10",
            "{{value_r11_c18}}": "R10",
            "{{value_r11_c19}}": "S10",
            "{{value_r11_c21}}": "U10",
            "{{value_r11_c22}}": "V10",
            "{{value_r11_c23}}": "W10",
            "{{value_r11_c24}}": "X10",
            "{{value_r11_c25}}": "Y10",
            "{{value_r11_c26}}": "Z10",
            
            # Row 12 (original row 11 in xlwings)  
            "{{value_r12_c2}}": "B11",
            "{{value_r12_c3}}": "C11",
            "{{value_r12_c4}}": "D11",
            "{{value_r12_c5}}": "E11",
            "{{value_r12_c6}}": "F11",
            "{{value_r12_c8}}": "H11",
            "{{value_r12_c9}}": "I11",
            "{{value_r12_c10}}": "J11",
            "{{value_r12_c11}}": "K11",
            "{{value_r12_c12}}": "L11",
            "{{value_r12_c14}}": "N11",
            "{{value_r12_c15}}": "O11",
            "{{value_r12_c16}}": "P11",
            "{{value_r12_c17}}": "Q11",
            "{{value_r12_c18}}": "R11",
            "{{value_r12_c19}}": "S11",
            "{{value_r12_c21}}": "U11",
            "{{value_r12_c22}}": "V11",
            "{{value_r12_c23}}": "W11",
            "{{value_r12_c24}}": "X11",
            "{{value_r12_c25}}": "Y11",
            "{{value_r12_c26}}": "Z11",
        }
        
        # Generate HTML using Google Sheets data
        success = sheets_manager.generate_html_from_sheet(
            spreadsheet_id=SPREADSHEET_ID,
            template_path=template_path,
            output_path=output_file,
            cell_mappings=cell_mappings
        )
        
        if success:
            print(f"✅ Successfully generated HTML report: {output_file}")
            
            # Upload the generated HTML to Google Drive
            reports_folder_id = path_manager.get_drive_folder_id("root")  # Adjust as needed
            if reports_folder_id:
                drive_manager.upload_file(output_file, reports_folder_id, f"Daily_Pickup_Summary_{PROPERTY_CODE}_{current_date}.html")
                print(f"✅ Uploaded HTML report to Google Drive")
        else:
            print(f"❌ Failed to generate HTML report")
            
        return success
        
    except Exception as e:
        print(f"❌ An error occurred generating BWOF HTML: {e}")
        return False

def main():
    """Main execution function"""
    print("BWOF HTML Generation Script - Google Sheets API Version")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)
    
    success = generate_bwof_html()
    
    print("="*60) 
    if success:
        print("✅ BWOF HTML generation completed successfully")
    else:
        print("❌ BWOF HTML generation failed")
        print("📝 Note: Make sure SPREADSHEET_ID is configured with actual Google Sheets ID")
        
    print(f"Finished: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

if __name__ == "__main__":
    main()
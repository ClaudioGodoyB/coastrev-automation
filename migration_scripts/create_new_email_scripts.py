#!/usr/bin/env python3
"""
Create New Email Scripts for All Properties
Generates Gmail API-based email download scripts to replace win32com.client versions
"""
import os
from pathlib import Path

# Property configurations based on analysis of existing scripts
PROPERTY_CONFIGS = {
    "BWOF": {"type": "outlook", "subject": "BWOF Daily Report"},
    "BWSM": {"type": "outlook", "subject": "BWSM Daily Report"}, 
    "CHA": {"type": "outlook", "subject": "OTB Statistics - Carlton Hotel"},
    "CSI": {"type": "outlook", "subject": "CSI Daily Report"},
    "HATH": {"type": "cloudbeds", "subject": "rooms sold and occupancy"},
    "HBSB": {"type": "cloudbeds", "subject": "rooms sold and occupancy"},
    "ISLO": {"type": "outlook", "subject": "ISLO Daily Report"},
    "LOL": {"type": "cloudbeds", "subject": "rooms sold and occupancy"},
    "LSMB": {"type": "outlook", "subject": "LSMB Daily Report"},
    "MPMS": {"type": "outlook", "subject": "MPMS Daily Report"},
    "MSI": {"type": "outlook", "subject": "MSI Daily Report"},
    "RRM": {"type": "outlook", "subject": "RRM Daily Report"},
    "TLMB": {"type": "cloudbeds", "subject": "rooms sold and occupancy"},
    "TRH": {"type": "cloudbeds", "subject": "rooms sold and occupancy"}
}

EMAIL_SCRIPT_TEMPLATE = '''#!/usr/bin/env python3
"""
{property_code} Email Download - VM Compatible Version  
Property: {property_code} 
Subject: '{subject_keywords}'
Type: {email_type}
"""
import os
import sys
from datetime import datetime

# Add utils to path for imports
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'utils'))
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'config'))

from gmail_manager import gmail_manager
from google_drive_manager import drive_manager
from paths import path_manager

# {property_code} CONFIGURATION
PROPERTY_CODE = "{property_code}"
SUBJECT_KEYWORDS = "{subject_keywords}"
EMAIL_TYPE = "{email_type}"

def download_{property_lower}_emails():
    """Download {property_code} property emails using Gmail API"""
    try:
        print(f"🏨 Processing emails for {property_code}")
        print(f"🔍 Subject keywords: {{SUBJECT_KEYWORDS}}")
        print(f"📧 Email type: {{EMAIL_TYPE}}")
        print("="*50)
        
        # Initialize managers
        if not gmail_manager.service:
            print("❌ Gmail service not available")
            return False
            
        if not drive_manager.service:
            print("❌ Google Drive service not available") 
            return False
        
        # Process emails based on type
        success = False
        
        if EMAIL_TYPE.lower() == "outlook":
            # Regular email attachment processing
            success = gmail_manager.process_outlook_attachments(
                property_code=PROPERTY_CODE,
                subject_keyword=SUBJECT_KEYWORDS
            )
            
        elif EMAIL_TYPE.lower() == "cloudbeds":
            # Cloudbeds link extraction and CSV download
            success = gmail_manager.process_cloudbeds_emails(
                property_code=PROPERTY_CODE,
                subject_filter=SUBJECT_KEYWORDS
            )
            
        else:
            print(f"❌ Unknown email type: {{EMAIL_TYPE}}")
            return False
        
        if success:
            print(f"✅ Successfully processed {property_code} emails")
        else:
            print(f"⚠️ No {property_code} emails found")
            
        return success
        
    except Exception as e:
        print(f"❌ An error occurred processing {property_code} emails: {{e}}")
        return False

def main():
    """Main execution function"""
    print("{property_code} Email Download Script - Gmail API Version")
    print(f"Started: {{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}}")
    print("="*60)
    
    success = download_{property_lower}_emails()
    
    print("="*60) 
    if success:
        print("✅ {property_code} email processing completed successfully")
    else:
        print("❌ {property_code} email processing failed or found no emails")
        
    print(f"Finished: {{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}}")

if __name__ == "__main__":
    main()
'''

def create_email_script(property_code, config, base_dir):
    """Create a new email script for a property"""
    try:
        # Determine the script directory
        script_dir = base_dir / f"Scripts - {property_code}"
        
        if not script_dir.exists():
            print(f"❌ Script directory not found: {script_dir}")
            return False
            
        # Determine filename based on type
        if config["type"] == "cloudbeds":
            filename = "0.1 - Cloudbeds Email Download Attachement NEW.py"
        else:
            filename = "0.1 - Email Download Attachement NEW.py"
            
        script_path = script_dir / filename
        
        # Generate script content
        script_content = EMAIL_SCRIPT_TEMPLATE.format(
            property_code=property_code,
            property_lower=property_code.lower(),
            subject_keywords=config["subject"],
            email_type=config["type"]
        )
        
        # Write the script
        with open(script_path, 'w', encoding='utf-8') as f:
            f.write(script_content)
            
        print(f"✅ Created: {script_path}")
        return True
        
    except Exception as e:
        print(f"❌ Failed to create script for {property_code}: {e}")
        return False

def main():
    """Main execution function"""
    print("Creating New Email Scripts for All Properties")
    print("="*60)
    
    # Get the base directory
    script_dir = Path(__file__).parent
    base_dir = script_dir.parent  # Go up one level from migration_scripts/
    
    created_count = 0
    failed_count = 0
    
    # Create scripts for all properties
    for property_code, config in PROPERTY_CONFIGS.items():
        print(f"Creating script for {property_code}...")
        
        if create_email_script(property_code, config, base_dir):
            created_count += 1
        else:
            failed_count += 1
    
    print("="*60)
    print("SCRIPT CREATION SUMMARY")
    print("="*60)
    print(f"Successfully created: {created_count} scripts")
    print(f"Failed to create: {failed_count} scripts")
    print(f"Total properties: {len(PROPERTY_CONFIGS)}")
    
    if created_count > 0:
        print(f"✅ New email scripts created successfully!")
        print("📝 Note: Review and test each script before deploying")
    else:
        print("❌ No email scripts were created")

if __name__ == "__main__":
    main()
#!/usr/bin/env python3
"""
BWSM Email Download - VM Compatible Version  
Property: BWSM 
Subject: 'BWSM Daily Report'
Type: outlook
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

# BWSM CONFIGURATION
PROPERTY_CODE = "BWSM"
SUBJECT_KEYWORDS = "BWSM Daily Report"
EMAIL_TYPE = "outlook"

def download_bwsm_emails():
    """Download BWSM property emails using Gmail API"""
    try:
        print(f"🏨 Processing emails for BWSM")
        print(f"🔍 Subject keywords: {SUBJECT_KEYWORDS}")
        print(f"📧 Email type: {EMAIL_TYPE}")
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
            print(f"❌ Unknown email type: {EMAIL_TYPE}")
            return False
        
        if success:
            print(f"✅ Successfully processed BWSM emails")
        else:
            print(f"⚠️ No BWSM emails found")
            
        return success
        
    except Exception as e:
        print(f"❌ An error occurred processing BWSM emails: {e}")
        return False

def main():
    """Main execution function"""
    print("BWSM Email Download Script - Gmail API Version")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)
    
    success = download_bwsm_emails()
    
    print("="*60) 
    if success:
        print("✅ BWSM email processing completed successfully")
    else:
        print("❌ BWSM email processing failed or found no emails")
        
    print(f"Finished: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

if __name__ == "__main__":
    main()

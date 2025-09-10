#!/usr/bin/env python3
"""
HATH Email Download - VM Compatible Version  
Property: HATH 
Subject: 'rooms sold and occupancy'
Type: cloudbeds
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

# HATH CONFIGURATION
PROPERTY_CODE = "HATH"
SUBJECT_KEYWORDS = "rooms sold and occupancy"
EMAIL_TYPE = "cloudbeds"

def download_hath_emails():
    """Download HATH property emails using Gmail API"""
    try:
        print(f"🏨 Processing emails for HATH")
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
            print(f"✅ Successfully processed HATH emails")
        else:
            print(f"⚠️ No HATH emails found")
            
        return success
        
    except Exception as e:
        print(f"❌ An error occurred processing HATH emails: {e}")
        return False

def main():
    """Main execution function"""
    print("HATH Email Download Script - Gmail API Version")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)
    
    success = download_hath_emails()
    
    print("="*60) 
    if success:
        print("✅ HATH email processing completed successfully")
    else:
        print("❌ HATH email processing failed or found no emails")
        
    print(f"Finished: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

if __name__ == "__main__":
    main()

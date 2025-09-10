#!/usr/bin/env python3
"""
Email Download Template - VM Compatible Version
Replace win32com.client with Gmail API calls

Usage: Copy this template and modify the PROPERTY_CODE and SUBJECT_KEYWORDS
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

# CONFIGURATION - MODIFY THESE FOR EACH PROPERTY
PROPERTY_CODE = "TEMPLATE"  # Change this: BWOF, CHA, HATH, etc.
SUBJECT_KEYWORDS = "template subject"  # Change this: email subject to search for
EMAIL_TYPE = "outlook"  # Change this: "outlook" for regular attachments, "cloudbeds" for link extraction

def download_property_emails():
    """Download emails for this property using Gmail API"""
    try:
        print(f"🏨 Processing emails for {PROPERTY_CODE}")
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
            print(f"✅ Successfully processed emails for {PROPERTY_CODE}")
        else:
            print(f"⚠️ No emails processed for {PROPERTY_CODE}")
            
        return success
        
    except Exception as e:
        print(f"❌ An error occurred processing {PROPERTY_CODE}: {e}")
        return False

def main():
    """Main execution function"""
    print(f"Email Download Script - {PROPERTY_CODE}")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)
    
    success = download_property_emails()
    
    print("="*60) 
    if success:
        print("✅ Email processing completed successfully")
    else:
        print("❌ Email processing failed or found no emails")
        
    print(f"Finished: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

if __name__ == "__main__":
    main()
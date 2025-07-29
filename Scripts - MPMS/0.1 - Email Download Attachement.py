import os
import win32com.client
from datetime import datetime

try:
    # Get today's date
    today = datetime.now().date()
    today_str = today.strftime("%Y-%m-%d")

    # Folder to save attachments
    target_dir = rf"C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Extracts\Extract {today_str}\MPMS"
    os.makedirs(target_dir, exist_ok=True)

    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Inbox
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    match_found = False

    for msg in messages:
        # Ensure it's from today and matches the subject case-insensitively
        if msg.ReceivedTime.date() == today and "mph5786 pace" in msg.Subject.lower():
            match_found = True
            attachments = msg.Attachments
            for i in range(1, attachments.Count + 1):
                attachment = attachments.Item(i)
                filename = attachment.FileName
                save_path = os.path.join(target_dir, filename)
                attachment.SaveAsFile(save_path)
                print(f"✅ Saved attachment: {save_path}")

    if not match_found:
        print("⚠️ No matching emails with subject containing 'MPH5786 Pace' were found today.")

except Exception as e:
    print(f"❌ An error occurred: {e}")

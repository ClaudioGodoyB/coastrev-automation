import os #Variable items = LSMB (Property nickname), 33076 (ASI Property ID)
import win32com.client
import datetime
from datetime import timedelta

def download_attachments(subject_keyword, save_folder, folder_date):
    # Updated folder structure: "Extract YYYY-MM-DD\LSMB"
    folder_name = f"Extract {folder_date}\\LSMB"
    full_folder_path = os.path.join(save_folder, folder_name)

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Check if the folder exists before proceeding
    if os.path.exists(full_folder_path):
        messages = outlook.GetDefaultFolder(6).Items  # 6 corresponds to the Inbox folder
        messages.Sort("[ReceivedTime]", True)  # Sort by ReceivedTime in descending order

        found_message = False
        for message in messages:
            if message.SenderEmailAddress == 'noreply@anandsystems.email':  # Check if the sender matches
                if folder_date in message.ReceivedTime.strftime('%Y-%m-%d') and subject_keyword in message.Subject:
                    found_message = True
                    for attachment in message.Attachments:
                        # Only process files containing '33076' in their name
                        if '33076' in attachment.FileName:
                            filename = os.path.join(full_folder_path, attachment.FileName)
                            if not os.path.exists(filename):
                                attachment.SaveAsFile(filename)
                                print(f"Attachment saved: {filename} in folder: '{folder_name}'")
                            else:
                                print(f"Skipped existing file: {filename} in folder: '{folder_name}'")
        
        if not found_message:
            print(f"No email found for {folder_date} from 'noreply@anandsystems.email' with subject containing '{subject_keyword}' in folder: '{folder_name}'.")
    else:
        print(f"Folder '{folder_name}' does not exist. Skipping download for this folder.")

def get_most_recent_weekend():
    today = datetime.datetime.now()
    # Find the most recent Saturday (0 = Monday, 6 = Sunday)
    last_saturday = today - timedelta(days=(today.weekday() + 2) % 7)
    # Find the most recent Sunday (0 = Monday, 6 = Sunday)
    last_sunday = today - timedelta(days=(today.weekday() + 1) % 7)

    return last_saturday.strftime("%Y-%m-%d"), last_sunday.strftime("%Y-%m-%d")

if __name__ == "__main__":
    subject_keyword = "Summary"  # Change this to your desired subject keyword
    save_folder = r"/home/user/coastrev/data/extracts"  # Change this to your desired save folder

    # Create today's folder and download today's attachments
    today_date = datetime.datetime.now().strftime("%Y-%m-%d")
    download_attachments(subject_keyword, save_folder, today_date)

    # Create folders and download attachments for the most recent Saturday and Sunday
    last_saturday, last_sunday = get_most_recent_weekend()
    download_attachments(subject_keyword, save_folder, last_saturday)
    download_attachments(subject_keyword, save_folder, last_sunday)

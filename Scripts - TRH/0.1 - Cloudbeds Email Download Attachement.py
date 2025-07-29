import os #Variables = TRH [Property nickname], line 31 [Target email Subject]
import win32com.client
import re
import requests
from datetime import datetime

def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "", name).strip().replace(" ", "_")

try:
    # Today's date
    today = datetime.now().date()
    today_str = today.strftime("%Y-%m-%d")

    # Build dynamic folder path
    base_dir = rf"C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Extracts\Extract {today_str}\TRH"
    os.makedirs(base_dir, exist_ok=True)

    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    seen_links = set()
    file_tracker = {}  # Tracks how many files per report name

    for msg in messages:
        if (msg.SenderEmailAddress.lower() == "noreply@cloudbeds.com" and
            msg.ReceivedTime.date() == today and
            msg.Subject.strip().lower() == "Rooms Sold and Occupancy - TRH".lower()):

            html = msg.HTMLBody

            # Extract report title
            title_match = re.search(r'<a[^>]*>([^<]+)</a>', html, re.IGNORECASE)
            report_title = title_match.group(1).strip() if title_match else "cloudbeds_report"
            clean_title = sanitize_filename(report_title)

            # Extract download link
            match = re.search(r'href="(https://link\.cloudbeds\.com/[^"]+)"', html)
            if match:
                download_link = match.group(1)

                if download_link in seen_links:
                    continue  # skip duplicates

                seen_links.add(download_link)

                response = requests.get(download_link)
                if response.status_code == 200:
                    # Determine filename versioning
                    base_name = f"{clean_title}_{today_str}"
                    count = file_tracker.get(base_name, 0) + 1
                    file_tracker[base_name] = count

                    filename = f"{base_name}_{count}.csv"
                    file_path = os.path.join(base_dir, filename)

                    with open(file_path, 'wb') as f:
                        f.write(response.content)
                    print(f"✅ Saved: {file_path}")
                else:
                    print(f"❌ Failed to download file from link:\n{download_link}")

    if not file_tracker:
        print("⚠️ No matching emails or valid links were found today.")

except Exception as e:
    print(f"❌ An error occurred in the script but continuing gracefully: {e}")

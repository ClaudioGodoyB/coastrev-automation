import os
import subprocess
from datetime import datetime

# Format today's date
today_str = datetime.today().strftime('%Y-%m-%d')
daily_detail_folder = f"Daily Detail {today_str}"
final_file_name = f"BWSM {today_str}.xlsx"
final_file_path = os.path.join(
    r"C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Details",
    daily_detail_folder,
    final_file_name
)

# Check if the final output file already exists
if os.path.exists(final_file_path):
    print(f"✅ Final file already exists: {final_file_path}. Skipping script.")
else:
    # Define the directory containing the scripts
    script_directory = r'C:\Users\johnj\Desktop\CoastRev\Reporting\Scripts\Scripts - BWSM'

    # Loop through all files in the directory
    for filename in os.listdir(script_directory):
        if filename.startswith("0.") and filename.endswith(".py"):
            script_path = os.path.join(script_directory, filename)
            print(f"🚀 Running script: {filename}")
            try:
                subprocess.run(["python", script_path], check=True)
            except subprocess.CalledProcessError as e:
                print(f"⚠️ Script {filename} failed with error: {e}. Continuing to next.")

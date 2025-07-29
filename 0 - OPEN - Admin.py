import os
import subprocess

# Define the directory containing the scripts
script_directory = r'C:\Users\johnj\Desktop\CoastRev\Reporting\Scripts\Scripts - Admin'

# Loop through all files in the directory
for filename in os.listdir(script_directory):
    if filename.startswith("0.") and filename.endswith(".py"):  # Check if the file starts with "0." and is a Python script
        script_path = os.path.join(script_directory, filename)
        print(f"Running script: {filename}")
        subprocess.run(["python", script_path], check=True)  # Execute the script

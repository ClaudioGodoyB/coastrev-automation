import os # No changing varianbles
from datetime import datetime

def create_folder_with_date(path, subfolders=[]):
    # Get today's date in the format YYYY-MM-DD
    today_date = datetime.now().strftime("%Y-%m-%d")

    # Combine the given path with the folder name "Daily Detail" and today's date
    new_folder_path = os.path.join(path, f"Daily Detail {today_date}")

    try:
        # Create the new folder
        os.mkdir(new_folder_path)
        print(f"Folder 'Daily Detail {today_date}' created successfully at {new_folder_path}")

        # Create subfolders within the newly created folder
        for subfolder in subfolders:
            subfolder_path = os.path.join(new_folder_path, subfolder)
            os.mkdir(subfolder_path)
            print(f"Subfolder '{subfolder}' created successfully at {subfolder_path}")

    except OSError as e:
        print(f"Error: {e}")

# Example usage:
given_path = "C:\\Users\\johnj\\Desktop\\CoastRev\\Reporting\\Daily Details"  # Replace with your desired path

# Define subfolder names
subfolders = ['Misc']  # You can adjust this list to add or remove subfolders

# Create today's folder with subfolders
create_folder_with_date(given_path, subfolders)

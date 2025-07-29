import os #Variables at line 38 with list of property IDs
from datetime import datetime, timedelta

def create_folder_with_date(path, folder_date, subfolders=[]):
    # Combine the given path with the folder name "Extract" and the folder_date
    new_folder_path = os.path.join(path, f"Extract {folder_date}")

    # Check if the folder already exists
    if not os.path.exists(new_folder_path):
        try:
            # Create the new folder
            os.mkdir(new_folder_path)
            print(f"Folder 'Extract {folder_date}' created successfully at {new_folder_path}")
            
            # Create subfolders within the newly created folder
            for subfolder in subfolders:
                subfolder_path = os.path.join(new_folder_path, subfolder)
                os.mkdir(subfolder_path)
                print(f"Subfolder '{subfolder}' created successfully at {subfolder_path}")
        except OSError as e:
            print(f"Error: {e}")
    else:
        print(f"Folder 'Extract {folder_date}' already exists.")

def get_most_recent_weekend():
    today = datetime.now()
    # Find the most recent Saturday (0 = Monday, 6 = Sunday)
    last_saturday = today - timedelta(days=(today.weekday() + 2) % 7)
    # Find the most recent Sunday (0 = Monday, 6 = Sunday)
    last_sunday = today - timedelta(days=(today.weekday() + 1) % 7)

    return last_saturday.strftime("%Y-%m-%d"), last_sunday.strftime("%Y-%m-%d")

# Example usage:
given_path = "C:\\Users\\johnj\\Desktop\\CoastRev\\Reporting\\Daily Extracts"  # Replace with your desired path

# Define subfolder names
subfolders = ['LOL', 'TRH', 'TLMB', 'RRM', 'ISLO', 'HBSB', 'HATH', 'MSI', 'BWSM', 'CHA', 'LSMB']  # You can adjust this list to add or remove subfolders

# Create today's folder with subfolders
today_date = datetime.now().strftime("%Y-%m-%d")
create_folder_with_date(given_path, today_date, subfolders)

# Create folders for the most recent Saturday and Sunday with subfolders
last_saturday, last_sunday = get_most_recent_weekend()
create_folder_with_date(given_path, last_saturday, subfolders)
create_folder_with_date(given_path, last_sunday, subfolders)

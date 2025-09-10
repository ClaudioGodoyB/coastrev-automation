import os # No changing variables
from openpyxl import load_workbook
from datetime import datetime

def repair_and_save_excel(source_file, destination_folder):
    # Check if the file exists
    if not os.path.exists(source_file):
        print(f"File not found: {source_file}")
        return
    
    try:
        # Load the workbook
        print(f"Processing file: {source_file}")
        workbook = load_workbook(source_file)

        # Ensure the destination folder exists
        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)

        # Define the destination file path
        file_name = os.path.basename(source_file)
        destination_file = os.path.join(destination_folder, file_name)

        # Save the workbook (repaired if possible) to the destination folder
        workbook.save(destination_file)
        print(f"Repaired and saved: {destination_file}")

    except Exception as e:
        print(f"Failed to open or repair the file: {source_file}\nError: {e}")

def find_and_process_files(downloads_folder, save_root_folder):
    # Get today's date in YYYY_MM_DD format (to match your file format)
    today = datetime.today().strftime('%Y_%m_%d')

    # Filter for files containing 'expedia_price_grid' and today's date in the filename
    excel_files = [f for f in os.listdir(downloads_folder) 
                   if f.endswith('.xlsx') and 'expedia_price_grid' in f and today in f]

    if not excel_files:
        print("No matching files found for today.")
        return

    # Create the folder name with today's date in YYYY-MM-DD format
    destination_folder = os.path.join(save_root_folder, f'Extract {datetime.today().strftime("%Y-%m-%d")}')

    # Process each file
    for excel_file in excel_files:
        source_file = os.path.join(downloads_folder, excel_file)
        repair_and_save_excel(source_file, destination_folder)

if __name__ == '__main__':
    # Define the paths
    downloads_folder = r"/home/user/coastrev/data/downloads"
    save_root_folder = r"/home/user/coastrev/data/extracts"

    # Find today's downloaded 'expedia_price_grid' files and process them
    find_and_process_files(downloads_folder, save_root_folder)

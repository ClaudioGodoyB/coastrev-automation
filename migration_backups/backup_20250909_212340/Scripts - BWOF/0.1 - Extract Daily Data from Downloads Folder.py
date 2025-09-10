import os #Variables = BWOF [Property Nickname], 10437 [Property Brand ID]
import shutil
from datetime import datetime

def is_file_from_today(file_path):
    file_date = datetime.fromtimestamp(os.path.getctime(file_path))
    return file_date.date() == datetime.today().date()

def copy_csv_file(source_file, destination_folder):
    try:
        print(f"Copying file: {source_file}")

        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)

        file_name = os.path.basename(source_file)
        destination_file = os.path.join(destination_folder, file_name)

        shutil.copy2(source_file, destination_file)
        print(f"Copied to: {destination_file}")

    except Exception as e:
        print(f"Failed to copy the file: {source_file}\nError: {e}")

def find_and_copy_file(downloads_folder, save_root_folder):
    csv_files = [f for f in os.listdir(downloads_folder)
                 if f.lower().endswith('.csv') and '10437_metrics' in f.lower()]

    if not csv_files:
        print("No matching CSV files found.")
        return

    dated_folder = f'Extract {datetime.today().strftime("%Y-%m-%d")}'
    destination_folder = os.path.join(save_root_folder, dated_folder, 'BWOF')

    copied_any = False
    for csv_file in csv_files:
        source_file = os.path.join(downloads_folder, csv_file)
        if is_file_from_today(source_file):
            copy_csv_file(source_file, destination_folder)
            copied_any = True

    if not copied_any:
        print("No matching files found from today.")

if __name__ == '__main__':
    downloads_folder = r'C:\Users\johnj\Downloads'
    save_root_folder = r'C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Extracts'

    find_and_copy_file(downloads_folder, save_root_folder)

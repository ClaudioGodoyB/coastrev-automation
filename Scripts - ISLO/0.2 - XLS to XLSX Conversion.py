import os #Variable items = ISLO (Property nickname)
import xlrd
from openpyxl import Workbook
from datetime import datetime, timedelta

def convert_xls_to_xlsx(directory):
    if not os.path.exists(directory):
        print(f"Directory '{directory}' does not exist. Skipping conversion for this folder.")
        return

    files_converted = 0
    files_skipped = 0

    # Loop through all the files in the directory
    for filename in os.listdir(directory):
        if filename.endswith(".xls") and "Forecasting" in filename:
            # Full file path
            file_path = os.path.join(directory, filename)

            # Load the .xls file using xlrd
            xls_file = xlrd.open_workbook(file_path)
            sheet = xls_file.sheet_by_index(0)

            # Create a new .xlsx workbook
            workbook = Workbook()
            new_sheet = workbook.active

            # Loop through the rows and columns of the .xls file and write them to the .xlsx file
            for row in range(sheet.nrows):
                for col in range(sheet.ncols):
                    new_sheet.cell(row=row + 1, column=col + 1).value = sheet.cell_value(row, col)

            # Define the new file name with .xlsx extension
            new_filename = filename.replace(".xls", ".xlsx")
            new_file_path = os.path.join(directory, new_filename)

            # Save the .xlsx file
            if not os.path.exists(new_file_path):
                workbook.save(new_file_path)
                files_converted += 1
                print(f"Converted {filename} to {new_filename} in folder: {directory}")
            else:
                files_skipped += 1
                print(f"Skipped existing file: {new_filename} in folder: {directory}")

    if files_converted == 0 and files_skipped == 0:
        print(f"Folder '{directory}' is empty or contains no valid .xls files to convert.")
    elif files_converted == 0 and files_skipped > 0:
        print(f"All .xls files in folder '{directory}' have already been converted.")
    else:
        print(f"Conversion completed for folder '{directory}': {files_converted} file(s) converted, {files_skipped} file(s) skipped.")

def get_most_recent_weekend():
    today = datetime.now()
    # Find the most recent Saturday (0 = Monday, 6 = Sunday)
    last_saturday = today - timedelta(days=(today.weekday() + 2) % 7)
    # Find the most recent Sunday (0 = Monday, 6 = Sunday)
    last_sunday = today - timedelta(days=(today.weekday() + 1) % 7)

    return last_saturday.strftime("%Y-%m-%d"), last_sunday.strftime("%Y-%m-%d")

if __name__ == "__main__":
    # Get the current date in the desired format
    current_date = datetime.now().strftime("%Y-%m-%d")

    # Define the base directory path
    base_directory = r"/home/user/coastrev/data/extracts"

    # Define the folder path with the current date and the 'ISLO' subfolder
    today_directory = os.path.join(base_directory, f"Extract {current_date}", "ISLO")

    # Perform conversion for today's 'ISLO' subfolder
    convert_xls_to_xlsx(today_directory)

    # Get the folders for the most recent Saturday and Sunday
    last_saturday, last_sunday = get_most_recent_weekend()

    # Define paths for 'ISLO' subfolder of the latest Saturday and Sunday folders
    saturday_directory = os.path.join(base_directory, f"Extract {last_saturday}", "ISLO")
    sunday_directory = os.path.join(base_directory, f"Extract {last_sunday}", "ISLO")

    # Perform conversion for the latest Saturday and Sunday 'ISLO' subfolders
    convert_xls_to_xlsx(saturday_directory)
    convert_xls_to_xlsx(sunday_directory)

    print("Conversion process completed.")

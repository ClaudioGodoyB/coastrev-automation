import xlwings as xw #Variables = LOL [Property Nickname]
import os
from datetime import datetime

try:
    # File path to your destination file
    file_path = r"C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Templates\LOL.xlsx"

    # Generate the destination folder path based on the current date
    current_date = datetime.now().strftime("%Y-%m-%d")
    destination_folder = os.path.join(r"C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Details", f"Daily Detail {current_date}")

    # Ensure the destination folder exists
    os.makedirs(destination_folder, exist_ok=True)

    # Construct the new file path
    new_file_path = os.path.join(destination_folder, f"LOL {current_date}.xlsx")

    # Open the workbook
    app = xw.App(visible=False)  # You can set visible=True if you want to see the Excel window
    wb = xw.Book(file_path)

    # Activate the 'Summary' tab
    ws = wb.sheets['Summary']
    ws.activate()

    # Select cell A1
    ws.range('A1').select()

    # Save the workbook to the new path
    wb.save(new_file_path)

    # Close the workbook
    wb.close()

    # Quit the app
    app.quit()

    # If everything runs without errors, print Success
    print("Success")
except Exception as e:
    # If an error occurs, print Failed and the error message
    print(f"Failed: {e}")

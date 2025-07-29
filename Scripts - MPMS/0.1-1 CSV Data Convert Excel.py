import os #Variables = MPMS [Property Nickname]
import xlwings as xw
import csv
from datetime import date

# Get today's date and construct dynamic folder path
today_str = date.today().strftime('%Y-%m-%d')
folder_path = fr"C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Extracts\Extract {today_str}\MPMS"

# Process all .csv files in the folder
for filename in os.listdir(folder_path):
    if filename.lower().endswith(".csv"):
        file_path = os.path.join(folder_path, filename)
        base_name = os.path.splitext(filename)[0]
        excel_file_path = os.path.join(folder_path, f"{base_name}.xlsx")

        # Read the CSV file
        with open(file_path, newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            data = list(reader)

        # Write to Excel
        with xw.App(visible=False) as app:
            wb = app.books.add()
            ws = wb.sheets[0]
            ws.range("A1").value = data
            wb.save(excel_file_path)
            wb.close()

        print(f"Converted and saved: {excel_file_path}")

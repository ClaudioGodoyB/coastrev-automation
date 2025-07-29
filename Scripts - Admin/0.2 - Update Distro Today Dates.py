import xlwings as xw # No changing variables

try:
    # Define the file path
    file_path = r"C:\Users\johnj\Desktop\CoastRev\Reporting\Daily Templates\!Daily Detail Roster & Distro.xlsx"

    # Open the Excel file
    wb = xw.Book(file_path)

    # Refresh all the formulas in the workbook
    wb.app.calculate()

    # Save the changes
    wb.save()

    # Close the workbook
    wb.close()

    # If everything runs without errors, print Success
    print("Success")
except Exception as e:
    # If an error occurs, print Failed and the error message
    print(f"Failed: {e}")

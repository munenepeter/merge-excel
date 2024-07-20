import os
import win32com.client as client

# Define paths
old_version_path = os.path.join(os.getcwd(), "old version")
new_version_path = os.path.join(os.getcwd(), "new version")

# Ensure new version directory exists
if not os.path.exists(new_version_path):
    os.makedirs(new_version_path)

# Initialize Excel application
excel = client.Dispatch("Excel.Application")

try:
    for file in os.listdir(old_version_path):
        filename, fileextension = os.path.splitext(file)
        if fileextension.lower() in ['.xls', '.xlsx']:  # Add any other Excel extensions if needed
            wb = excel.Workbooks.Open(os.path.join(old_version_path, file))
            output_path = os.path.join(new_version_path, filename + ".xlsx")
            wb.SaveAs(output_path, FileFormat=51)  # 51 represents the XLSX format
            wb.Close(False)
finally:
    excel.Quit()

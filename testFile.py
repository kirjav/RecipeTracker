import openpyxl
from openpyxl import Workbook
import os

# Function to create a new Excel file with headers if it doesn't exist
def create_excel_file(file_path):
    try:
        # Try opening an existing file
        wb = openpyxl.load_workbook(file_path)
        print(f"Workbook '{file_path}' loaded.")
    except FileNotFoundError:
        # If file does not exist, create a new one
        wb = Workbook()
        sheet = wb.active
        # Define headers for the food tracking sheet
        sheet.append(["Date", "Recipe Link", "Dish Name", "Cooking Method", "Notes"])
        wb.save(file_path)
        print(f"Workbook '{file_path}' created.")
    
    return wb


if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    filename = "food_log.xlsx"  # Excel file name
    file_path = os.path.join(script_dir, filename)
    create_excel_file(file_path)
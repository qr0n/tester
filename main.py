from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image
import pyautogui
import pygetwindow as gw
import time
from extractor import *

c_file = "F:/projects/compsci/mod5.c"

workbook_path = 'C:/Users/bhave/OneDrive/Documents/test.xlsx'

workbook = load_workbook(workbook_path)
alphabet_dict = {i: chr(i + ord('A') - 1) for i in range(1, 27)}

sheet = workbook['Sheet1']
left, top, width, height = 100, 100, 500, 500

class ExcelFileManagement:
    @staticmethod
    def add_value(cell_id, text):
        try:
            sheet[cell_id] = text
            workbook.save(workbook_path)
            return "Done"
        except Exception as E:
                return E
        
    @staticmethod
    def add_image_to_cell(cell_id, image_path, width, height):
        try:
            img = Image(image_path)
            img.width = width
            img.height = height
            sheet.add_image(img, cell_id)
            workbook.save(workbook_path)
            return "Done"
        except Exception as E:
            return E
            
    @staticmethod
    def populate_table():
        headers = ["Test number", "Test Type", "Variables", "Module", "Input", "Expected result", "Actual result", "Screenshot"]
    # Calculate the starting cell for populating data
        if sheet.max_row == 1:
            start_row = 1
        else:
            start_row = sheet.max_row + 5
        start_cell = f"{alphabet_dict[1]}{start_row}"

        for i, header in enumerate(headers, start=1):
            sheet.cell(row=i + start_row - 1, column=1, value=header)

        print(f"Starting from cell: {start_cell}")
        workbook.save(workbook_path)
        
    @staticmethod
    def add_results(test_type, variables, module, input_data, expected):
        try:
        # Find the next available row
            if sheet.max_row == 1:
                row = sheet.max_row
            else:
                row = sheet.max_row - 7
        
        # Define the starting column index
            start_col = 2  # Column B
        
        # Set values in the corresponding cells
            sheet.cell(row=row + 1, column=start_col).value = test_type
            sheet.cell(row=row + 2, column=start_col).value = variables
            
            sheet.cell(row=row + 3, column=start_col).value = module
            sheet.cell(row=row + 4, column=start_col).value = input_data
            
            sheet.cell(row=row + 5, column=start_col).value = "Placeholder for standard out"
            sheet.cell(row=row + 6, column=start_col).value = "Placeholder for sceenshot"

        # Save the workbook
            workbook.save(workbook_path)
            return "Results added successfully"
        except Exception as e:
            return str(e)

class ScreenshotManagement:
    class Helper:
        def focus_window_by_title(window_title):
            try:
                window = gw.getWindowsWithTitle(window_title)[0]
                window.activate()
                return True
            except IndexError:
                print(f"Window with title '{window_title}' not found.")
                return False
            
    @staticmethod
    def take_screenshot(path_to_save):
        ScreenshotManagement.Helper.focus_window_by_title("Windows PowerShell")
        time.sleep(1)
        screenshot = pyautogui.screenshot(region=(left, top, width, height))
        screenshot.save(path_to_save)
    
# ScreenshotManagement.take_screenshot("C:/Users/bhave/OneDrive/Desktop/test.png")
# ExcelFileManagement.populate_table()
# ExcelFileManagement.add_results("main", "hello data", "nothing yet")

class UserManagement:
    class Helper:
        @staticmethod
        def prompt():
            type_input = input("What type of test are you running? (Normal/Erroneous/Extreme/Incomplete)?\n> ")
            print(f"Reading file located at {c_file}")
            print("Extracting function signatures...")
            extract(c_file)
            module_input = input("Select the function you're testing from the list above (or if it is not there enter a new one)\n> ")
            variable_input = input("Paste the variables your code is using here\n> ")
            input_input = input("Enter the input data:\n> ")
            expected_input = input("What do you expect this code to return?\n> ")

            result = ExcelFileManagement.add_results(module=module_input, input_data=input_input, expected=expected_input)
            if result == "Results added successfully":
                print("Results added successfully!")
            else:
                print(f"Error: {result}")

# Example usage:
# print(alphabet_dict[sheet.max_column], sheet.max_row)
ExcelFileManagement.populate_table()
ExcelFileManagement.add_results(test_type="normal", variables="1", module="2", input_data="3", expected="4")

"""
TODO: Replace placeholders for the real
TODO: Connect to STDOUT
"""

import tkinter as tk
import openpyxl
from tkinter import filedialog

file_path = ''


# Select the Excel file
def select_file():
    global file_path
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename()

    print(file_path)


# Open the Excel file
def open_file():
    global file_path
    # Open the Excel file
    workbook = openpyxl.load_workbook(file_path, data_only=True)

    # Get the sheet you want to work with
    sheet = workbook['DUT']

    # Now you can access individual cells or ranges of cells
    # For example, to read the value of cell A1:
    dut_value = sheet['Y162'].value
    ref_value = sheet['AA162'].value

    if dut_value - ref_value > 5:
        print("FAIL")
    else:
        print("PASS")


def main():
    select_file()
    open_file()


if __name__ == "__main__":
    main()
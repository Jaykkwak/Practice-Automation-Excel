import tkinter as tk
import openpyxl
from tkinter import filedialog

from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.workbook import Workbook

global file_path
global sheet


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
    global sheet
    # Open the Excel file
    workbook = openpyxl.load_workbook(file_path, data_only=True)

    # Get the sheet you want to work with
    sheet = workbook['DUT']


def get_values():
    global sheet
    # Now you can access individual cells or ranges of cells
    # For example, to read the value of cell A1:
    dut_value = sheet['Y162'].value
    ref_value = sheet['AA162'].value

    # check the value
    if dut_value - ref_value > 5:
        print("FAIL")
    else:
        print("PASS")


def copy_cells():
    global sheet

    dest_wb = Workbook()
    dest_ws = dest_wb.active

    # Define the range to copy
    copy_range = sheet['D16:M23']

    top_left_cell = 'A1'

    # Calculate the column offset directly
    column_offset = column_index_from_string(top_left_cell[0]) - 1

    # Calculate the row offset
    row_offset = int(top_left_cell[1:]) - 1

    # Iterate over the source range
    for row in copy_range:
        for cell in row:
            # Calculate the destination cell coordinates
            dest_cell = dest_ws.cell(row=cell.row + row_offset - 1, column=cell.column + column_offset - 1)
            # Copy value and style from the source cell to the destination cell
            dest_cell.value = cell.value

    dest_wb.save('destination.xlsx')


def main():
    select_file()
    open_file()
    # get_values()
    copy_cells()


if __name__ == "__main__":
    main()

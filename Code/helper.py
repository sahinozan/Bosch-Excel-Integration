# Author: Ozan Şahin

# TODO: Delete print statements
# TODO: Integrate input validation with the UI

import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl.styles import PatternFill, Font
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from importlib.util import find_spec
from subprocess import check_call, check_output
from sys import exit
import os
import sys
from transformer_ui import show_error
from path_checker import path_validation


# This will not be needed when the script is converted to a standalone executable
def package_control(packages: list):
    for package in packages:
        if find_spec(package) is None:
            print(f"\n>>> Installing {package}...\n")
            check_call([sys.executable, '-m', 'pip', 'install', package,
                        '--disable-pip-version-check'])


package_control(packages=["pandas", "openpyxl", "numpy"])

# import modules after checking if they exist in the environment

# Colors for Excel formatting
redFill = PatternFill(start_color='FFFF0000',
                      end_color='FFFF0000',
                      fill_type='solid')


# separators for different operating systems are used to make it compatible with all of them
# directory separator = "/" for Linux and MacOS
# directory separator = "\\" for Windows

def file_path_handler():
    # validation added to fix VSCode issue (will be removed later)
    if str(os.getcwd()).split(os.sep)[-1] == "Code":
        directory = check_output(["python", f"{os.getcwd()}{os.sep}transformer_ui.py"])
    else:
        directory = check_output(["python", f"{os.getcwd()}{os.sep}Code{os.sep}transformer_ui.py"])
    directory = directory.decode("utf-8")
    directory = str(directory.strip())
    paths = {}

    # Create a dictionary from the output of the UI
    # The dictionary will be used to extract the file paths
    for i in directory.split(os.linesep):
        component = i.split("=")
        if len(component) > 1:
            paths[component[0]] = component[1]

    sorted_paths = sorted(paths.items(), key=lambda item: item[0], reverse=True)
    paths = {key: path for key, path in sorted_paths}

    # input validation for file paths
    path_validation(paths)

    # create output directory name from the source file path
    current_source_dir, past_source_dir, output_dir = paths["Source1"], paths["Source2"], paths["Output"]
    current_source_file_name = current_source_dir.split("/")[-1]
    output_dir = output_dir + os.sep + current_source_file_name.split(".")[0] + "_output.xlsx"

    try:
        current_source_file = pd.read_excel(current_source_dir)
        past_source_file = pd.read_excel(past_source_dir)
    except FileNotFoundError:
        show_error("File not found!")
        exit(0)

    # Validation will not be needed after the standalone executable
    if str(os.getcwd()).split(os.sep)[-1] == "Code":
        master_path = os.sep.join(str(os.getcwd()).split(os.sep)[:-1]) + \
                      f"{os.sep}Data{os.sep}Master Data.xlsx"
    else:
        master_path = os.getcwd() + f"{os.sep}Data{os.sep}Master Data.xlsx"
    try:
        pipes = pd.read_excel(master_path, sheet_name="Cihaz - Boru - Miktar")
        types = pd.read_excel(master_path, sheet_name="Boru - Tip")
    except FileNotFoundError:
        show_error("File not found!")
        exit(0)

    return current_source_file, past_source_file, pipes, types, output_dir


# General formatting such as column width, row height, and coloring
# Adds auto-filter to every column except the first one (index column)
def general_excel_formatter(file_path: str) -> None:
    wb = openpyxl.load_workbook(file_path)

    ws1 = wb["Sheet1"]
    ws1.delete_rows(3)

    dim_holder = DimensionHolder(worksheet=ws1)

    for col in range(ws1.min_column, ws1.max_column + 1):
        dim_holder[get_column_letter(col)] = ColumnDimension(ws1, min=col, max=col, width=12)

    #  change the height of all rows
    for row in range(ws1.min_row, ws1.max_row + 1):
        ws1.row_dimensions[row].height = 20

    # change the size of B, C, and D columns
    dim_holder['B'] = ColumnDimension(ws1, min=2, max=2, width=18)
    dim_holder['C'] = ColumnDimension(ws1, min=3, max=3, width=18)
    dim_holder['D'] = ColumnDimension(ws1, min=4, max=4, width=18)
    dim_holder['E'] = ColumnDimension(ws1, min=5, max=5, width=18)

    # add filter
    ws1.auto_filter.ref = "A2:AC2"

    # highlight the version and date cells
    ws1['A1'].fill = redFill
    ws1['A1'].font = Font(color="FFFFFF", bold=True, size=11)

    ws1['B1'].fill = redFill
    ws1['B1'].font = Font(color="FFFFFF", bold=True, size=11)

    ws1.column_dimensions = dim_holder
    wb.save(file_path)


# similar to general_excel_formatter but for the pivoted sheet
def pivot_excel_formatter(file_path: str) -> None:
    wb = openpyxl.load_workbook(file_path)

    ws2 = wb["Sheet2"]

    ws2.delete_rows(3)

    dim_holder = DimensionHolder(worksheet=ws2)

    for col in range(ws2.min_column, ws2.max_column + 1):
        dim_holder[get_column_letter(col)] = ColumnDimension(ws2, min=col, max=col, width=12)

    #  change the height of all rows
    for row in range(ws2.min_row, ws2.max_row + 1):
        ws2.row_dimensions[row].height = 20

    # change the size of B, C, and D columns
    dim_holder['B'] = ColumnDimension(ws2, min=2, max=2, width=18)
    dim_holder['C'] = ColumnDimension(ws2, min=3, max=3, width=18)

    # add filter
    ws2.auto_filter.ref = "A2:AA2"

    # highlight the version and date cells
    ws2['A1'].fill = redFill
    ws2['A1'].font = Font(color="FFFFFF", bold=True, size=11)

    ws2['B1'].fill = redFill
    ws2['B1'].font = Font(color="FFFFFF", bold=True, size=11)

    ws2.column_dimensions = dim_holder
    wb.save(file_path)


# Adds the version and date information to the Excel file
def excel_version(file_path: str, file: pd.DataFrame) -> None:
    wb = openpyxl.load_workbook(file_path)

    version = file.iloc[[3, 4], 7]
    update_date = file.iloc[[3, 4], 8]

    version_value = version.iloc[0] + " - " + version.iloc[1]
    update_date_value = update_date.iloc[0] + ":  " + update_date.iloc[1]

    for sheet in wb.sheetnames:
        ws = wb[sheet]
        ws.cell(row=1, column=1).value = version_value
        ws.cell(row=1, column=2).value = str(update_date_value)
        ws.cell(row=2, column=1).value = "Hat"

        # center all cells
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')

        wb.save(file_path)


# checks if the Excel file exists, if it does not then it creates it
def check_and_create_sheet(output_excel_file: str) -> None:
    try:
        if os.path.exists(output_excel_file):
            wb = openpyxl.load_workbook(output_excel_file)

            if "Sheet1" not in wb.sheetnames:
                wb.create_sheet("Sheet1")
            if "Sheet2" not in wb.sheetnames:
                wb.create_sheet("Sheet2")

            wb.save(output_excel_file)
    except PermissionError:
        show_error("Permission Error!")
        exit(1)
    except Exception as e:  # catch all other exceptions
        show_error(f"{e}!")
        exit(1)


# writes the dataframes to the two sheets in the Excel file (general-pivoted)
def write_to_excel(output_excel_file, main: pd.DataFrame, pivot: pd.DataFrame,
                   main_sheet_name="Sheet1", pivot_sheet_name="Sheet2") -> None:
    print(">>> Conversion started...")
    try:
        with pd.ExcelWriter(output_excel_file, mode="w") as writer:
            main.to_excel(writer, main_sheet_name)
            pivot.to_excel(writer, pivot_sheet_name)
        print(">>> Conversion completed!")
    except PermissionError:
        show_error("Conversion Failed!")
        show_error("Permission Error!")
        exit(1)
    except Exception as e:  # catch all other exceptions
        show_error(f"{e}!")
        exit(1)

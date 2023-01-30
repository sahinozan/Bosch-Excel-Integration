# Author: Ozan Sahin
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.styles import PatternFill, Font
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
import openpyxl
import pandas as pd
from subprocess import Popen, PIPE
from sys import exit
import os
import sys


def resource_path(relative_path, type="c"):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS  # type: ignore
    except Exception:
        if type == "d":
            base_path = os.environ.get("_MEIPASS2", os.path.abspath(f"..{os.sep}Data"))
        else:
            base_path = os.environ.get("_MEIPASS2", os.path.abspath("."))

    return os.path.join(base_path, relative_path)


# Colors for Excel formatting
redFill = PatternFill(start_color='FFFF0000',
                      end_color='FFFF0000',
                      fill_type='solid')


def file_path_handler():

    process = Popen(["python", f"{resource_path('transformer_ui.py')}"], stdout=PIPE, stderr=PIPE)
    _, _ = process.communicate()

    with open(resource_path("paths.txt"), "r") as file:
        directory = file.read()
        source_dir = directory.split(";")[0]
        output_dir = directory.split(";")[1]

    source_file_name = source_dir.split("/")[-1]
    output_dir = output_dir + os.sep + source_file_name.split(".")[0] + "_output.xlsx"

    try:
        source_file = pd.read_excel(source_dir)
    except FileNotFoundError:
        exit(0)

    pipes_path = resource_path("Cihazlar - Borular.xlsx", type="d")
    types_path = resource_path("Borular - Tipler.xlsx", type="d")
    
    try:
        pipes = pd.read_excel(pipes_path)
        types = pd.read_excel(types_path)
    except FileNotFoundError:
        exit(0)

    return source_file, pipes, types, output_dir


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
    ws1.auto_filter.ref = "A2:Z2"

    # highlight the version and date cells
    ws1['A1'].fill = redFill
    ws1['A1'].font = Font(color="FFFFFF", bold=True, size=11)

    ws1['B1'].fill = redFill
    ws1['B1'].font = Font(color="FFFFFF", bold=True, size=11)

    ws1.column_dimensions = dim_holder
    wb.save(file_path)


# similar to general_excel_formatter but for the pivotted sheet
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
    ws2.auto_filter.ref = "A2:Z2"

    # highlight the version and date cells
    ws2['A1'].fill = redFill
    ws2['A1'].font = Font(color="FFFFFF", bold=True, size=11)

    ws2['B1'].fill = redFill
    ws2['B1'].font = Font(color="FFFFFF", bold=True, size=11)

    ws2.column_dimensions = dim_holder
    wb.save(file_path)


# Adds the version and date information to the excel file
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


# checks if the excel file exists, if it does not then it creates it
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
        exit(1)
    except Exception as e:  # catch all other exceptions
        exit(1)


# writes the dataframes to the two sheets in the excel file (general-pivotted)
def write_to_excel(output_excel_file, main: pd.DataFrame, pivot: pd.DataFrame,
                   main_sheet_name="Sheet1", pivot_sheet_name="Sheet2") -> None:
    try:
        with pd.ExcelWriter(output_excel_file, mode="w") as writer:
            main.to_excel(writer, main_sheet_name)
            pivot.to_excel(writer, pivot_sheet_name)
    except PermissionError:
        exit(1)
    except Exception as e:  # catch all other exceptions
        exit(1)

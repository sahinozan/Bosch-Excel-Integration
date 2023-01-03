from importlib.util import find_spec
from subprocess import check_call, check_output
from sys import exit
import os
import sys

def package_control(packages: list):
    for package in packages:
        if find_spec(package) is None:
            print(f"\n>>> Installing {package}...\n")
            check_call([sys.executable, '-m', 'pip', 'install', package,
                       '--disable-pip-version-check'])


package_control(packages=["pandas", "openpyxl", "numpy"])

from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.styles import PatternFill, Font
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
import openpyxl
import pandas as pd

# Colors for Excel formatting
redFill = PatternFill(start_color='FFFF0000',
                      end_color='FFFF0000',
                      fill_type='solid')

def file_path_handler(user_input: str):
    if user_input == "Y":
        excel_file = input(">>> Enter the Excel file name: ")

        if str(os.getcwd()).split("/")[-1] == "Code":
            source_path = "/".join(str(os.getcwd()).split("/")[:-1]) + \
                "/Data/Source/" + excel_file + ".xlsx"
            try:
                file = pd.read_excel(source_path)
            except FileNotFoundError:
                print("!!> File not found!")
                exit(0)
            output_excel_file = "/".join(str(os.getcwd()).split("/")[:-1]) + \
                "/Data/Output/" + excel_file + "_output.xlsx"
        else:
            source_path = os.getcwd() + "/Data/Source/" + excel_file + ".xlsx"
            try:
                file = pd.read_excel(os.getcwd() +
                                     "/Data/Source/" + excel_file + ".xlsx")
            except FileNotFoundError:
                print("!!> File not found!")
                exit(0)
            output_excel_file = os.getcwd() + "/Data/Output/" + excel_file + "_output.xlsx"
    elif user_input == "N":
        if str(os.getcwd()).split("/")[-1] == "Code":
            source_file = check_output(["python", f"{os.getcwd()}/transformer_ui.py"])
        else:
            source_file = check_output(["python", f"{os.getcwd()}/Code/transformer_ui.py"])
        source_file = source_file.decode("utf-8")
        source_file = str(source_file.strip())

        if len(source_file) == 0:
            print("!!> You have not selected an Excel file!")
            exit(0)

        output_excel_file_dir = source_file.split("/")[:-1]
        output_excel_file_name = list(source_file.split("/")[-1].split(".")[0])
        output_excel_file_dir[output_excel_file_dir.index("Source")] = "Output"
        output_excel_file_dir = "/".join(output_excel_file_dir)
        output_excel_file_name.append("_output.xlsx")
        output_excel_file_name = "".join(output_excel_file_name)
        output_excel_file_dir = "".join(output_excel_file_dir)
        output_excel_file = os.path.join(output_excel_file_dir, output_excel_file_name)

        try:
            file = pd.read_excel(source_file)
        except FileNotFoundError:
            print("!!> File not found!")
            exit(0)
    else:
        raise SystemError("!!> Invalid input!")

    if str(os.getcwd()).split("/")[-1] == "Code":
        pipes_path = "/".join(str(os.getcwd()).split("/")[:-1]) + \
            "/Data/Cihazlar - Borular.xlsx"
        types_path = "/".join(str(os.getcwd()).split("/")[:-1]) + \
            "/Data/Borular - Tipler.xlsx"
    else:
        pipes_path = os.getcwd() + "/Data/Cihazlar - Borular.xlsx"
        types_path = os.getcwd() + "/Data/Borular - Tipler.xlsx"
    try:
        pipes = pd.read_excel(pipes_path)
        types = pd.read_excel(types_path)
    except FileNotFoundError:
        print("!!> File not found!")
        exit(0)

    return file, pipes, types, output_excel_file


def python_version_control() -> None:
    if sys.version_info.major != 3:
        raise SystemError("!!> Python version must be 3.X.X (preferably, 3.11.X)")
    if sys.version_info.minor != 11:
        print("??> This script is written for Python 3.11.X, "
              "It may not work properly with other versions.\n"
              "??> Do you still want to continue? (Y/N)")
        user_input = input(">>> ")
        if user_input == 'Y':
            pass
        elif user_input == 'N':
            print(">>> Terminating...")
            exit(0)
        else:
            raise SystemError("!!> Invalid input!")


def general_excel_formatter(file_path: str) -> None:
    wb = openpyxl.load_workbook(file_path)

    ws1 = wb["Sheet1"]  # type: Worksheet
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
    ws1.auto_filter.ref = "A2:E2"

    # highlight the version and date cells
    ws1['A1'].fill = redFill
    ws1['A1'].font = Font(color="FFFFFF", bold=True, size=11)

    ws1['B1'].fill = redFill
    ws1['B1'].font = Font(color="FFFFFF", bold=True, size=11)

    ws1.column_dimensions = dim_holder
    wb.save(file_path)


def pivot_excel_formatter(file_path: str) -> None:
    wb = openpyxl.load_workbook(file_path)

    ws2 = wb["Sheet2"]  # type: Worksheet

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
    ws2.auto_filter.ref = "A2:C2"

    # highlight the version and date cells
    ws2['A1'].fill = redFill
    ws2['A1'].font = Font(color="FFFFFF", bold=True, size=11)

    ws2['B1'].fill = redFill
    ws2['B1'].font = Font(color="FFFFFF", bold=True, size=11)

    ws2.column_dimensions = dim_holder
    wb.save(file_path)


def excel_version(file_path: str, version: str, update_date: str) -> None:
    wb = openpyxl.load_workbook(file_path)

    for sheet in wb.sheetnames:
        ws = wb[sheet]  # type: Worksheet
        ws.cell(row=1, column=1).value = version
        ws.cell(row=1, column=2).value = str(update_date)
        ws.cell(row=2, column=1).value = "Hat"

        # center all cells
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')

        wb.save(file_path)


def check_and_create_sheet(output_excel_file: str) -> None:

    print(">>>\n>>> Validating Excel Files...")
    try:
        if os.path.exists(output_excel_file):
            wb = openpyxl.load_workbook(output_excel_file)

            if "Sheet1" not in wb.sheetnames:
                wb.create_sheet("Sheet1")
            if "Sheet2" not in wb.sheetnames:
                wb.create_sheet("Sheet2")

            wb.save(output_excel_file)

        print(">>> Validation completed!")
    except PermissionError:
        print("!!> Permission denied!")
        print(">>> Terminating...")
        exit(1)
    except Exception as e:  # catch all other exceptions
        print(f"!!> {e}")
        print(">>> Terminating...")
        exit(1)


def write_to_excel(output_excel_file, main: pd.DataFrame, pivot: pd.DataFrame,
                   main_sheet_name="Sheet1", pivot_sheet_name="Sheet2") -> None:

    print(">>>\n>>> Conversion started...")

    try:
        with pd.ExcelWriter(output_excel_file, mode="w") as writer:
            main.to_excel(writer, main_sheet_name)
            pivot.to_excel(writer, pivot_sheet_name)

        print(">>> Conversion completed!")
    except PermissionError:
        print("!!> Conversion failed!")
        print("!!> Permission denied!")
        print(">>> Terminating...")
        exit(1)
    except Exception as e:  # catch all other exceptions
        print(f"!!> {e}")
        print(">>> Terminating...")
        exit(1)

from __future__ import annotations
from importlib.util import find_spec
import warnings
import subprocess
import datetime
import sys

warnings.filterwarnings("ignore")

# check compatibility with Python 3.11
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

#  check if required packages are installed
if find_spec('pandas') is None:
    print("\n>>> Installing pandas...\n")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pandas', '--disable-pip-version-check'])
if find_spec("openpyxl") is None:
    print("\n>>> Installing openpyxl...\n")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl', '--disable-pip-version-check'])

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder, RowDimension
from openpyxl.styles import PatternFill, Font
from openpyxl.styles import Alignment
import openpyxl
import pandas as pd

# read data source files
try:
    file = pd.read_excel('../Data/KW47_V00.xlsx')
    pipes = pd.read_excel('../Data/Cihazlar - Borular.xlsx')
except FileNotFoundError:
    print("File not found!")
    exit(1)

shift_date = file.iloc[4:6, 12: 31: 3].copy()
shift_date.iloc[0, :] = shift_date.iloc[0, :].apply(lambda x: x.strftime("%d %b %Y"))
shift_date = shift_date.apply(lambda x: f"{x.iloc[1]} - {x.iloc[0]}", axis=0)
shift_dates = list(shift_date)

version = file.iloc[[3, 4], 7]
update_date = file.iloc[[3, 4], 8]

indices = file.iloc[:, [0, 7, 8, 11]].reset_index()
work_days = file.iloc[:, 12: 33].reset_index()
sheet = pd.concat([indices, work_days], axis=1).iloc[2:, :]

sheet = sheet[sheet.iloc[:, 1].notna() & sheet.iloc[:, 2].notna()]
sheet.iloc[:, 2] = sheet.iloc[:, 2].astype(str)
sheet = sheet[sheet.iloc[:, 2].apply(str.isnumeric)]

sheet = sheet[sheet.iloc[:, 6].apply(lambda x: (type(x) != datetime.datetime) and (type(x) != str))]
sheet = sheet[sheet.iloc[:, 3].notna()]

sheet.drop('index', axis=1, inplace=True)
sheet.reset_index(drop=True, inplace=True)

initial_indices = ["Hat", "Cihaz TTNr", "Cihaz Aile", "Toplam Adet"]
shifts = [1, 2, 3]
final_indices = [" - ".join([i, str(j)]) for i in shift_dates for j in shifts]
initial_indices.extend(final_indices)

sheet = sheet.set_axis(initial_indices, axis=1)

sheet["Cihaz TTNr"] = sheet["Cihaz TTNr"].astype(str)
pipes["Cihaz"] = pipes["Cihaz"].astype(str)

sheet = sheet.merge(pipes, left_on="Cihaz TTNr", right_on="Cihaz", how="inner")
sheet.drop("Cihaz", axis=1, inplace=True)
sheet.insert(3, 'Boru', sheet.pop('Boru'))
sheet["Cihaz TTNr"] = sheet["Cihaz TTNr"].astype("int64")

indices = ["Hat", "Cihaz TTNr", "Cihaz Aile", "Boru TTNr"]

df = pd.DataFrame(columns=pd.MultiIndex.from_product([shift_dates, shifts]),
                  index=range(sheet.shape[0]))
df = pd.concat([pd.DataFrame(columns=pd.MultiIndex.from_product([indices,
                                                                 ["" for _ in range(len(indices))]])), df], axis=1)
dates_df = df.iloc[:, 16:]
initial_df = df.swaplevel(axis=1, i=0, j=1).iloc[:, :16]
initial_df = initial_df.loc[:, ~initial_df.columns.duplicated()]  # type: ignore
df = pd.concat([initial_df, dates_df], axis=1)

sheet.iloc[:, 4:] = sheet.iloc[:, 4:].apply(lambda x: x * sheet.iloc[:, -1], axis=0)

sheet.drop("Miktar", axis=1, inplace=True)
sheet.drop("Toplam Adet", axis=1, inplace=True)

df.iloc[:, :] = sheet.iloc[:, :]

df = df.set_index(("", "Hat")).rename_axis(None, axis=0)

# Colors for excel formatting
redFill = PatternFill(start_color='FFFF0000',
                   end_color='FFFF0000',
                   fill_type='solid')


# format the Excel column dimensions
def excel_formatter(file_path: str):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    ws.delete_rows(3)

    dim_holder = DimensionHolder(worksheet=ws)

    for col in range(ws.min_column, ws.max_column + 1):
        dim_holder[get_column_letter(col)] = ColumnDimension(ws, min=col, max=col, width=12)

    # change the height of all rows
    for row in range(ws.min_row, ws.max_row + 1):
        ws.row_dimensions[row].height = 20

    # change the size of B, C, and D columns
    dim_holder['B'] = ColumnDimension(ws, min=2, max=2, width=18)
    dim_holder['C'] = ColumnDimension(ws, min=3, max=3, width=18)
    dim_holder['D'] = ColumnDimension(ws, min=4, max=4, width=18)

    # add filter to the first column
    ws.auto_filter.ref = "A2:D2"

    # highlight the version and date cells
    ws['A1'].fill = redFill
    ws['A1'].font = Font(color = "FFFFFF", bold=True, size=11)

    ws['B1'].fill = redFill
    ws['B1'].font = Font(color = "FFFFFF", bold=True, size=11)

    ws.column_dimensions = dim_holder
    wb.save(file_path)


version_value = version.iloc[0] + " - " + version.iloc[1]
update_date_value = update_date.iloc[0] + ":  " + update_date.iloc[1]


# add the Excel version to the file
def excel_version(file_path: str):
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    ws.cell(row=1, column=1).value = version_value
    ws.cell(row=1, column=2).value = str(update_date_value)   # type: ignore
    ws.cell(row=2, column=1).value = "Hat"  # type: ignore

    # center all cells 
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')
    
    wb.save(file_path)


# write the dataframe to an Excel file
try:
    print(">>>\n>>> Conversion started...")
    df.to_excel("../Data/source.xlsx")
    print(">>> Conversion completed successfully!")
    print(">>> Excel Formatting started...")
    excel_formatter(file_path="../Data/source.xlsx")
    excel_version(file_path="../Data/source.xlsx")
    print(">>> Excel Formatting completed successfully!")
except PermissionError:
    print(">>> Conversion failed!")
finally:
    print(">>> Terminating...")
    exit(0)

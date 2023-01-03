from __future__ import annotations
from Utility import *
from sys import exit
import warnings
from subprocess import check_output
import datetime

warnings.filterwarnings("ignore")

# control if the required Python version is installed
python_version_control()

# control if the required packages are installed and install them if not
package_control(packages=["pandas", "openpyxl", "numpy", "tkinter"])

import numpy as np
import pandas as pd

# get user input and read data source files
print(">>> Do you want to manually enter the Excel file name? (Y/N)")
user_input = input(">>> ")
file, pipes, types, output_excel_file = file_path_handler(user_input=user_input)

#  get the date index range
date_start_index = str(file.columns[file.isin(['Pazartesi']).any()][0]).split(' ')[1]

shift_date = file.iloc[4:6, int(date_start_index): 31: 3].copy()
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
sheet = sheet.merge(types, on="Boru", how="left")
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

sheet.iloc[:, 4:-1] = sheet.iloc[:, 4:-1].apply(lambda x: x * sheet.iloc[:, -2], axis=0)

sheet.drop("Miktar", axis=1, inplace=True)
sheet.drop("Toplam Adet", axis=1, inplace=True)

df.iloc[:, :] = sheet.iloc[:, :]
df["Tip"] = sheet["Tip"]

type_df = df.swaplevel(axis=1, i=0, j=1).iloc[:, -1]
df = pd.concat([df.iloc[:, :-1], type_df], axis=1)
df.insert(4, ('', 'Tip'), df.pop(('', 'Tip')))  # type: ignore

df = df.set_index(("", "Hat")).rename_axis(None, axis=0)

df.iloc[:, 4:] = df.iloc[:, 4:].apply(pd.to_numeric, errors='coerce')

df_pivot = df.sort_index(key=lambda x: (x.to_series().str[6:].astype("int64")))
df_pivot = df_pivot.drop(columns=[('', 'Cihaz TTNr'), ('', 'Cihaz Aile'), ('', 'Tip')])

df_pivot.index = df_pivot.index.str.split(' ').str[1]
df_pivot = df_pivot.groupby([df_pivot.index, ("", "Boru TTNr")]).sum().sort_index(ascending=False)
df_pivot = df_pivot.reset_index(level=1, drop=False)
df_pivot.index = df_pivot.index.map(lambda x: f"Yalın {x}")
df_pivot["Tip"] = df_pivot.loc[:, ("", "Boru TTNr")].map(types.set_index("Boru")["Tip"])  # type: ignore

swapped_types = df_pivot.iloc[:, -1].to_frame().swaplevel(axis=1, i=0, j=1).iloc[:, 0].to_frame()
df_pivot.insert(1, ('', 'Tip'), swapped_types)  # type: ignore
df_pivot.drop(("Tip", ""), axis=1, inplace=True)

df_pivot.iloc[:, 2:] = df_pivot.iloc[:, 2:].applymap(lambda x: np.nan if x == 0 else x)

version_value = version.iloc[0] + " - " + version.iloc[1]
update_date_value = update_date.iloc[0] + ":  " + update_date.iloc[1]

# check if the sheets exist in the Excel file and create them if they don't
check_and_create_sheet(output_excel_file)

# write the dataframe to an Excel file
write_to_excel(output_excel_file, main=df, pivot=df_pivot)

try:
    print(">>>\n>>> Excel Formatting started...")
    pivot_excel_formatter(file_path=output_excel_file)
    general_excel_formatter(file_path=output_excel_file)
    excel_version(file_path=output_excel_file, version=version_value,
                  update_date=update_date_value)
    print(">>> Excel Formatting completed successfully!")
except PermissionError:
    print(">>> Formatting failed!")
finally:
    print(">>> Terminating...")
    exit(0)

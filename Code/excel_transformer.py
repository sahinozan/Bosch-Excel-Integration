from Utility import *
from sys import exit
import warnings
import datetime

warnings.filterwarnings("ignore")

# control if the required packages are installed and install them if not
package_control(packages=["pandas", "openpyxl", "numpy", "tkinter"])

# import modules after checking if they exist in the environment
import numpy as np
import pandas as pd

# read data source files
current_week, past_week, pipes, types, output_excel_file = file_path_handler()

# Disabled for now!
# check if the Excel file is in the desired format 
# TODO: create a more robust format control mechanism
# if len(first_file.columns[first_file.isin(['Pazartesi']).any()]) == 0:
#     print(">>> Format of the first Excel file is not desired. Use an appropriate formatted Excel file.")
#     exit(0)
# elif len(second_file.columns[first_file.isin(['Pazartesi']).any()]) == 0:
#     print(">>> Format of the second Excel file is not desired. Use an appropriate formatted Excel file.")
#     exit(0)

current_week = current_week.iloc[:, : 24]
past_week = pd.concat([past_week.iloc[:, :12], past_week.iloc[:, 21: 33]], axis=1)
master_file = pd.merge(current_week, past_week, how="left", left_index=True, right_index=True, suffixes=("_1", "_2"))

master_file.drop(master_file.iloc[:, 24:36], inplace=True, axis=1)
master_file_columns = list(master_file.columns)
master_file_columns = master_file_columns[:12] + master_file_columns[24:36] + master_file_columns[12:24]
master_file = master_file[master_file_columns]
first_file = master_file.copy()

# get the shift dates and format them (e.g. 27 Dec 2022)
shift_date = first_file.iloc[4:6, 12: 37: 3].copy()
shift_date.iloc[0, :] = shift_date.iloc[0, :].apply(lambda x: x.strftime("%d %b %Y"))
shift_date = shift_date.apply(lambda x: f"{x.iloc[1]} - {x.iloc[0]}", axis=0)
shift_dates = list(shift_date)

# TTNr, Hat, Cihaz Aile, and work days columns
indices = first_file.iloc[:, [0, 7, 8, 11]].reset_index()
work_days = first_file.iloc[:, 12: 36].reset_index()
sheet = pd.concat([indices, work_days], axis=1).iloc[2:, :]

# drop the rows with NaN values in the TTNr column
sheet = sheet[sheet.iloc[:, 1].notna() & sheet.iloc[:, 2].notna()]
sheet.iloc[:, 2] = sheet.iloc[:, 2].astype(str)
sheet = sheet[sheet.iloc[:, 2].apply(str.isnumeric)]

# get only the numeric values
sheet = sheet[sheet.iloc[:, 6].apply(lambda x: (type(x) != datetime.datetime) and (type(x) != str))]
sheet = sheet[sheet.iloc[:, 3].notna()]

# drop the index column (unnecessary)
sheet.drop('index', axis=1, inplace=True)
sheet.reset_index(drop=True, inplace=True)

# create the final dataframe indices
initial_indices = ["Hat", "Cihaz TTNr", "Cihaz Aile", "Toplam Adet"]
shifts = [1, 2, 3]
final_indices = [" - ".join([i, str(j)]) for i in shift_dates for j in shifts]
initial_indices.extend(final_indices)
sheet = sheet.set_axis(initial_indices, axis=1)

# convert to string to for the merge operation
sheet["Cihaz TTNr"] = sheet["Cihaz TTNr"].astype(str)
pipes["Cihaz"] = pipes["Cihaz"].astype(str)

# merge the pipes and devices dataframes 
sheet = sheet.merge(pipes, left_on="Cihaz TTNr", right_on="Cihaz", how="inner")
sheet = sheet.merge(types, on="Boru", how="left")

# drop the unnecessary columns
sheet.drop("Cihaz", axis=1, inplace=True)
sheet.insert(3, 'Boru', sheet.pop('Boru'))

# convert the Cihaz TTNr column back to int64 for Excel to have auto-filter on it
sheet["Cihaz TTNr"] = sheet["Cihaz TTNr"].astype("int64")
indices = ["Hat", "Cihaz TTNr", "Cihaz Aile", "Boru TTNr"]

# create the final dataframe with multi-level columns (same format with the initial Excel)
df = pd.DataFrame(columns=pd.MultiIndex.from_product([shift_dates, shifts]),
                  index=range(sheet.shape[0]))
df = pd.concat([pd.DataFrame(columns=pd.MultiIndex.from_product(
    [indices, ["" for _ in range(len(indices))]])), df], axis=1)

# swap levels according to the initial Excel and drop the duplicated columns
dates_df = df.iloc[:, 16:]
initial_df = df.swaplevel(axis=1, i=0, j=1).iloc[:, :16]
initial_df = initial_df.loc[:, ~initial_df.columns.duplicated()]
df = pd.concat([initial_df, dates_df], axis=1)
sheet.iloc[:, 4:-1] = sheet.iloc[:, 4:-1].apply(lambda x: x * sheet.iloc[:, -2], axis=0)

# drop the unnecessary columns
sheet.drop("Miktar", axis=1, inplace=True)
sheet.drop("Toplam Adet", axis=1, inplace=True)

# add the Tip column to the dataframe for the merge operation
df.iloc[:, :] = sheet.iloc[:, :]
df["Tip"] = sheet["Tip"]

# merge the dataframe with the types dataframe (hydraulic, spare, etc.)
type_df = df.swaplevel(axis=1, i=0, j=1).iloc[:, -1]
df = pd.concat([df.iloc[:, :-1], type_df], axis=1)
df.insert(4, ('', 'Tip'), df.pop(('', 'Tip')))

# dropped the index column name (will be filled later with openpyxl for better visuals)
df = df.set_index(("", "Hat")).rename_axis(axis=0)

# convert work days columns to numeric values 
df.iloc[:, 4:] = df.iloc[:, 4:].apply(pd.to_numeric, errors='coerce')

# create the dataframe for the Excel pivoting (multi-level columns)
df_pivot = df.sort_index(key=lambda x: (x.to_series().str[6:].astype("int64")))
df_pivot = df_pivot.drop(columns=[('', 'Cihaz TTNr'), ('', 'Cihaz Aile'), ('', 'Tip')])

# sum the values for the same pipe and shift (pivoting)
df_pivot.index = df_pivot.index.str.split(' ').str[1]
df_pivot = df_pivot.groupby([df_pivot.index, ("", "Boru TTNr")]).sum().sort_index(ascending=False)
df_pivot = df_pivot.reset_index(level=1, drop=False)
df_pivot.index = df_pivot.index.map(lambda x: f"Yalın {x}")
df_pivot["Tip"] = df_pivot.loc[:, ("", "Boru TTNr")].map(types.set_index("Boru")["Tip"])

# swap the levels of the columns to match the initial Excel
swapped_types = df_pivot.iloc[:, -1].to_frame().swaplevel(axis=1, i=0, j=1).iloc[:, 0].to_frame()
df_pivot.insert(1, ('', 'Tip'), swapped_types)
df_pivot.drop(("Tip", ""), axis=1, inplace=True)

# convert nan values to 0 to prevent errors in Excel
df_pivot.iloc[:, 2:] = df_pivot.iloc[:, 2:].applymap(lambda x: np.nan if x == 0 else x)

# check if the sheets exist in the Excel file and create them if they don't
check_and_create_sheet(output_excel_file)

# write the dataframe to an Excel file
write_to_excel(output_excel_file, main=df, pivot=df_pivot)

try:
    # progress bar will be added!
    # progress_bar.config(text="Excel Formatting started")
    pivot_excel_formatter(file_path=output_excel_file)
    general_excel_formatter(file_path=output_excel_file)
    excel_version(file_path=output_excel_file, file=first_file)
    # progress_bar.config(text="Excel Formatting completed successfully")
except PermissionError:
    show_error("Formatting Failed!")
finally:
    exit(0)

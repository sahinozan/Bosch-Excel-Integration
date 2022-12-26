from __future__ import annotations
from importlib.util import find_spec
import subprocess
import datetime
import sys

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

if find_spec('pandas') is None:
    print("\n>>> Installing pandas...\n")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pandas', '--disable-pip-version-check'])
if find_spec("openpyxl") is None:
    print("\n>>> Installing openpyxl...\n")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl', '--disable-pip-version-check'])
    
import pandas as pd

try:
    file = pd.read_excel('../Data/KW47_V00.xlsx')
    pipes = pd.read_excel('../Data/Cihazlar - Borular.xlsx')
except FileNotFoundError:
    print("File not found!")
    exit(1)

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

week_days = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
shifts = ["1", "2", "3"]

final_indices = [" ".join([i, j]) for i in week_days for j in shifts]
initial_indices.extend(final_indices)

sheet = sheet.set_axis(initial_indices, axis=1)
sheet.set_index("Hat")

sheet["Cihaz TTNr"] = sheet["Cihaz TTNr"].astype(str)
pipes["Cihaz"] = pipes["Cihaz"].astype(str)

sheet = sheet.merge(pipes, left_on="Cihaz TTNr", right_on="Cihaz", how="inner")
sheet.drop("Cihaz", axis=1, inplace=True)
sheet.insert(3, 'Boru', sheet.pop('Boru'))

for i in range(sheet.shape[0]):
    sheet.loc[i, "Pazartesi 1": "Pazar 3"] *= sheet.loc[i, "Miktar"]

sheet.drop("Miktar", axis=1, inplace=True)
sheet.drop("Toplam Adet", axis=1, inplace=True)
sheet.set_index("Hat", inplace=True)

try:
    sheet.to_excel("../Data/initial.xlsx")
    print("\n>>> Successfully converted to excel file!")
except PermissionError:
    print("\n>>> Failed to convert!")

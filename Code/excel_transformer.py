# Author: Ozan Şahin
import glob
from helper import *
from custom_ui import App
from customtkinter import CTk, filedialog

# def directory_selector():
#     root = CTk()
#     file_path = filedialog.askdirectory(parent=root, initialdir="/", title='Please select a directory')
#     root.destroy()
#     return file_path
#
#
# selected_dir = directory_selector()
# files_in_dir = glob.glob(pathname=selected_dir + "/*.xlsx")

# read data source files
source_file, pipes, types, output_excel_file, source_dir = file_path_handler()

# check whether the Excel files are in the desired format
excel_format_validate(list_of_dfs=[source_file])

# convert the dataframes into the desired format
source_file = general_excel_converter(raw_df=source_file, pipes=pipes, types=types)

# pivot the dataframe to eliminate duplicates
df_pivot = excel_pivoting(df_initial=source_file, types=types)

# check if the sheets exist in the Excel file and create them if they don't
check_and_create_sheet(output_excel_file=output_excel_file)

# write the dataframe to an Excel file
write_to_excel(output_excel_file=output_excel_file, main=source_file, pivot=df_pivot)

# format the Excel files
try:
    pivot_excel_formatter(file_path=output_excel_file)
    general_excel_formatter(file_path=output_excel_file, sheet_name="Genel")
    general_excel_formatter(file_path=output_excel_file, sheet_name="Pivot")
    excel_version(file_path=output_excel_file, file=source_file)
except PermissionError:
    App.show_error("Formatting Failed!")
finally:
    sys.exit(0)

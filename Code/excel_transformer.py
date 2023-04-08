# Author: Ozan Şahin
from helper import *
from custom_ui import App

# read data source files
source_file, pipes, types, output_excel_file, source_dir = file_path_handler()

# check whether the Excel files are in the desired format
excel_format_validate(list_of_dfs=[source_file])

# convert the dataframes into the desired format
source_df = general_excel_converter(raw_df=source_file, pipes=pipes, types=types)

# pivot the dataframe to eliminate duplicates
df_pivot = excel_pivoting(df_initial=source_df, types=types)

# check if the sheets exist in the Excel file and create them if they don't
check_and_create_sheet(output_excel_file=output_excel_file)

# write the dataframe to an Excel file
write_to_excel(output_excel_file=output_excel_file, main=source_df, pivot=df_pivot)

# format the Excel files
try:
    total_quantity_per_pipe(output_excel_file_path=output_excel_file)
    pivot_excel_formatter(file_path=output_excel_file)
    general_excel_formatter(file_path=output_excel_file, sheet_name="Genel")
    excel_version(file_path=output_excel_file, file=source_file)
except PermissionError:
    App.show_error("Formatting Failed!")
finally:
    sys.exit(0)

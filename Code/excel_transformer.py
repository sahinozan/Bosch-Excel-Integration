# Author: Ozan Şahin

from helper import *
from custom_ui import App
from rules import third_rule, shift_by_one

# read data source files
next_week, current_week, pipes, types, output_excel_file, current_dir, next_dir = file_path_handler()

# check whether the Excel files are in the desired format
excel_format_validate(list_of_dfs=[next_week, current_week])

# parse the source files and convert them into appropriate dataframes
next_week_df, current_week_df = source_file_parser(n_week_df=next_week, c_week_df=current_week)

# convert the dataframes into the desired format
current_week_df = general_excel_converter(raw_df=current_week_df, pipes=pipes, types=types)
next_week_df, initial_df, deleted_df = general_excel_converter(raw_df=next_week_df, pipes=pipes, types=types,
                                                               is_next_week=True)

# merge next week's data with current week's data
master_df = current_week_df.merge(next_week_df, on=[("", "Hat"), ("", "Cihaz TTNr"),
                                                    ("", "Cihaz Aile"), ("", "Boru TTNr"),
                                                    ("", "Tip")], how="right")

# detect devices without pipes
non_existing_df = detect_devices_without_pipes(source_df=initial_df, output_df=master_df)

# pivot the dataframe to eliminate duplicates
df_pivot = excel_pivoting(df_initial=master_df, types=types)

# check if the sheets exist in the Excel file and create them if they don't
check_and_create_sheet(output_excel_file=output_excel_file)

# write the dataframe to an Excel file
write_to_excel(output_excel_file=output_excel_file, main=master_df, pivot=df_pivot, non_existing=non_existing_df)

# format the Excel files
try:
    pivot_excel_formatter(file_path=output_excel_file)
    general_excel_formatter(file_path=output_excel_file, sheet_name="Genel")
    general_excel_formatter(file_path=output_excel_file, sheet_name="Borusuz")
    excel_version(file_path=output_excel_file, file=next_week)
except PermissionError:
    App.show_error("Formatting Failed!")
finally:
    sys.exit(0)

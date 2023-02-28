# Author: Ozan Şahin

import sys
import openpyxl
import pandas as pd
from customtkinter import CTkToplevel, CTkTabview, CTkLabel, CTkOptionMenu, CTkButton


class ShiftWindow(CTkToplevel):

    def __init__(self, *args, next_week_excel_path, current_week_excel_path, **kwargs):
        super().__init__(*args, **kwargs)

        self.dates = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        self.score_labels = []
        self.score_entries = []
        self.next_week_excel_path = next_week_excel_path
        self.current_week_excel_path = current_week_excel_path
        self.window_configuration()
        self.create_grid()

    def window_configuration(self) -> None:
        """
        Sets the title of the window, makes the window non-resizable, and makes the window close when the user clicks.
        Also sets the dimensions of the window and places it in the center of the screen.
        """
        self.title("Shift Selection")

        # make the window non-resizable (for macOS and Linux)
        if sys.platform == "darwin":
            self.resizable(False, False)

        # set the dimensions of the window
        w = 565  # width for the Tk root
        h = 420  # height for the Tk root

        # get screen width and height
        ws = self.winfo_screenwidth()  # width of the screen
        hs = self.winfo_screenheight()  # height of the screen

        # calculate x and y coordinates for the Tk root window
        x = (ws / 2) - (w / 2)
        y = (hs / 2) - (h / 2)

        # set the dimensions of the screen where the window will be displayed
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))

    def save_data(self):
        data = {'Gün': [date for date in self.dates]}
        df = pd.DataFrame(data)
        for shift in range(3):
            data[f'Shift {shift + 1}'] = [self.score_entries[self.dates.index(date) * 3 + shift].get() for date in
                                          self.dates]

        workbook = openpyxl.load_workbook(self.next_week_excel_path)
        worksheet = workbook['GZT-GWT']

        last_row = worksheet.max_row
        matching_rows = [i for i in range(1, last_row + 1) if
                         worksheet.cell(row=i, column=8).value and str(
                             worksheet.cell(row=i, column=8).value).startswith(
                             "7")]

        shifts = []

        for row in matching_rows:
            row_vardiya = []
            for col in range(13, 34):
                row_vardiya.append(worksheet.cell(row=row, column=col).value)
            shifts.append(row_vardiya)

        for i in range(len(shifts)):
            for j in range(len(shifts[i])):
                if shifts[i][j] is None:
                    shifts[i][j] = 0

        shift_list = []
        for i in range(len(df)):
            for j in range(len(df.columns)):
                if df.iloc[i, j] == "Var":
                    shift_list.append((i * 3) + j - 1)

        for k in range(len(shifts)):
            for i in shift_list:
                shifts[k][i - 1] = int(shifts[k][i]) + int(shifts[k][i - 1])
                shifts[k][i] = 0

        for row, lists in zip(matching_rows, shifts):
            for col, element in zip(range(13, 34), lists):
                worksheet.cell(row=row, column=col).value = element

        # DISABLED FOR DEBUGGING PURPOSES!
        # workbook.save(self.next_week_excel_path)

        self.destroy()

    def create_grid(self):
        current_week_tab_name = "Current Week"
        next_week_tab_name = "Next Week"

        current_week_identifier = self.current_week_excel_path.split("/")[-1].split(".")[0]
        next_week_identifier = self.next_week_excel_path.split("/")[-1].split(".")[0]

        tabview = CTkTabview(self)
        tabview.pack(fill="both", expand=True, pady=(0, 18))

        tabview.add(current_week_tab_name)
        tabview.add(next_week_tab_name)

        tab_names = [current_week_tab_name, next_week_tab_name]
        week_identifiers = [current_week_identifier, next_week_identifier]

        for index, tab_name in enumerate(tab_names):
            week_identifier = CTkLabel(tabview.tab(tab_name), text=week_identifiers[index])
            week_identifier.grid(row=0, column=0)

            header_labels = [CTkLabel(tabview.tab(tab_name), text=f"Shift {i + 1}") for i in range(3)]
            for i, header in enumerate(header_labels):
                header.grid(row=0, column=i + 1, sticky="nsew", padx=5, pady=5)

            for date in self.dates:
                date_label = CTkLabel(tabview.tab(tab_name), text=date)
                date_label.grid(row=self.dates.index(date) + 1, column=0, sticky="nsew", padx=15, pady=5)

                for shift in range(3):
                    score_entry = CTkOptionMenu(tabview.tab(tab_name), values=["Yok", "Var"], anchor="center")
                    score_entry.grid(row=self.dates.index(date) + 1, column=shift + 1, padx=5, pady=5, sticky="nsew")

                    self.score_entries.append(score_entry)
                    self.score_labels.append(date)

            save_button = CTkButton(tabview.tab(tab_name), text="Save", command=self.save_data)
            save_button.grid(row=len(self.dates) + 2, column=3, padx=5, pady=10)

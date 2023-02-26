import openpyxl
import tkinter as tk
from tkinter import ttk
import pandas as pd


def third_rule(input_excel_path):
    workbook = openpyxl.load_workbook(input_excel_path)
    worksheet = workbook['Pivot']

    last_row = worksheet.max_row
    matching_rows = [i for i in range(1, last_row + 1) if worksheet.cell(row=i, column=3).value and
                     (str(worksheet.cell(row=i, column=3).value).startswith("Hidrolik Borusu 2") or
                      "Hidrolik Borusu *" in str(worksheet.cell(row=i, column=3).value) or
                      "Yedek Parça *" in str(worksheet.cell(row=i, column=3).value))]

    vardiya = []

    for row in matching_rows:
        row_vardiya = []
        for col in range(4, 28):
            row_vardiya.append(worksheet.cell(row=row, column=col).value)
        vardiya.append(row_vardiya)

    for k in range(len(vardiya)):
        for i in range(1, 24):
            vardiya[k][i - 1] = vardiya[k][i]
            vardiya[k][i] = None

    for row, lists in zip(matching_rows, vardiya):
        for col, element in zip(range(4, 28), lists):
            worksheet.cell(row=row, column=col).value = element

    matching_rows2 = [i for i in range(1, last_row + 1) if worksheet.cell(row=i, column=1).value and
                      ("Yalın 3" in str(worksheet.cell(row=i, column=1).value) or
                       "Yalın 4" in str(worksheet.cell(row=i, column=1).value)) and
                      "Hidrolik Borusu" in str(worksheet.cell(row=i, column=3).value)]

    vardiya2 = []

    for row2 in matching_rows2:
        row_vardiya2 = []
        for col2 in range(4, 28):
            row_vardiya2.append(worksheet.cell(row=row2, column=col2).value)
        vardiya2.append(row_vardiya2)

    for k2 in range(len(vardiya2)):
        for i2 in range(1, 24):
            vardiya2[k2][i2 - 1] = vardiya2[k2][i2]
            vardiya2[k2][i2] = None

    for row2, lists2 in zip(matching_rows2, vardiya2):
        for col2, element2 in zip(range(4, 28), lists2):
            worksheet.cell(row=row2, column=col2).value = element2

    workbook.save(input_excel_path + '_rules.xlsx')


def first_rule(input_excel_path):
    new_file_extension = input_excel_path.split('/')[-1].replace('.xlsx', '_new.xlsx')
    new_input_excel_path = input_excel_path.replace(input_excel_path.split('/')[-1], new_file_extension)

    workbook = openpyxl.load_workbook(input_excel_path)
    worksheet = workbook['GZT-GWT']

    last_row = worksheet.max_row

    matching_rows = [i for i in range(1, last_row + 1) if worksheet.cell(row=i, column=1).value and
                     ("Yalın 7" in str(worksheet.cell(row=i, column=1).value) or
                      "Yalın 2" in str(worksheet.cell(row=i, column=1).value) or
                      "Yalın 4" in str(worksheet.cell(row=i, column=1).value) or
                      "Yalın 5" in str(worksheet.cell(row=i, column=1).value) or
                      "Yalın 6" in str(worksheet.cell(row=i, column=1).value) or
                      "Yalın 1" in str(worksheet.cell(row=i, column=1).value)) and
                     str(worksheet.cell(row=i, column=8).value).startswith("7")]

    vardiya = []

    for row in matching_rows:
        row_vardiya = []
        for col in range(13, 34):
            row_vardiya.append(worksheet.cell(row=row, column=col).value)
        vardiya.append(row_vardiya)

    for i in range(len(vardiya)):
        for j in range(len(vardiya[i])):
            if vardiya[i][j] is None:
                vardiya[i][j] = 0

    for k in range(len(vardiya)):
        for i in range(0, 20):
            if vardiya[k][i] < 70:
                vardiya[k][i - 1] = int(vardiya[k][i - 1]) + int(vardiya[k][i])
                vardiya[k][i] = 0
            else:
                vardiya[k][i] = int(vardiya[k][i]) - 40
                vardiya[k][i - 1] = int(vardiya[k][i - 1]) + 40

    for k in range(len(vardiya)):
        for i in range(0, 20):
            if 30 > vardiya[k][i] > 0 and vardiya[k][i + 1] > 0:
                vardiya[k][i + 1] = int(vardiya[k][i + 1]) + int(vardiya[k][i])
                vardiya[k][i] = 0

    for row, listeler in zip(matching_rows, vardiya):
        for col, eleman in zip(range(13, 34), listeler):
            worksheet.cell(row=row, column=col).value = eleman

    matching_rows2 = [i for i in range(1, last_row + 1) if worksheet.cell(row=i, column=1).value and
                      ("Yalın 3" in str(worksheet.cell(row=i, column=1).value)) and
                      str(worksheet.cell(row=i, column=8).value).startswith("7")]

    vardiya2 = []

    for row2 in matching_rows2:
        row_vardiya2 = []
        for col2 in range(13, 34):
            row_vardiya2.append(worksheet.cell(row=row2, column=col2).value)
        vardiya2.append(row_vardiya2)

    for i2 in range(len(vardiya2)):
        for j2 in range(len(vardiya2[i2])):
            if vardiya2[i2][j2] is None:
                vardiya2[i2][j2] = 0

    for k2 in range(len(vardiya2)):
        for i2 in range(0, 20):
            if vardiya2[k2][i2] < 60:
                vardiya2[k2][i2 - 1] = int(vardiya2[k2][i2 - 1]) + int(vardiya2[k2][i2])
                vardiya2[k2][i2] = 0
            else:
                vardiya2[k2][i2] = int(vardiya2[k2][i2]) - 30
                vardiya2[k2][i2 - 1] = int(vardiya2[k2][i2 - 1]) + 30

    for k2 in range(len(vardiya2)):
        for i2 in range(0, 20):
            if 30 > vardiya2[k2][i2] > 0 and vardiya2[k2][i2 + 1] > 0:
                vardiya2[k2][i2 + 1] = int(vardiya2[k2][i2 + 1]) + int(vardiya2[k2][i2])
                vardiya2[k2][i2] = 0

    for row2, listeler2 in zip(matching_rows2, vardiya2):
        for col2, eleman2 in zip(range(13, 34), listeler2):
            worksheet.cell(row=row2, column=col2).value = eleman2

    workbook.save(new_input_excel_path)


def second_rule(input_excel_path):
    new_file_extension = input_excel_path.split('/')[-1].replace('.xlsx', '_new.xlsx')
    new_input_excel_path = input_excel_path.replace(input_excel_path.split('/')[-1], new_file_extension)

    def save_data():
        data = {}
        data['Gün'] = [date for date in dates]
        for shift in range(3):
            data[f'Vardiya {shift + 1}'] = [score_entries[dates.index(date) * 3 + shift].get() for date in dates]
        global df
        df = pd.DataFrame(data)
        root.destroy()
        root.quit()

    root = tk.Tk()
    root.title("Mesai Planlama")

    dates = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]

    score_labels = []
    score_entries = []

    header_label = ttk.Label(root, text="Tarih")
    header_label.grid(row=0, column=0)
    header_labels = [ttk.Label(root, text=f"Vardiya {i + 1}") for i in range(3)]
    for i, header in enumerate(header_labels):
        header.grid(row=0, column=i + 1)

    for date in dates:
        date_label = ttk.Label(root, text=date)
        date_label.grid(row=dates.index(date) + 1, column=0)

        for shift in range(3):
            default_score = "Yok"

            score_entry = ttk.Entry(root)
            score_entry.insert(0, default_score)
            score_entry.grid(row=dates.index(date) + 1, column=shift + 1)

            score_entries.append(score_entry)
            score_labels.append(date)

    save_button = ttk.Button(root, text="Kaydet", command=save_data)
    save_button.grid(row=len(dates) + 2, column=0)

    root.mainloop()

    workbook = openpyxl.load_workbook(input_excel_path)
    worksheet = workbook['GZT-GWT']

    last_row = worksheet.max_row
    matching_rows = [i for i in range(1, last_row + 1) if
                     worksheet.cell(row=i, column=8).value and str(worksheet.cell(row=i, column=8).value).startswith(
                         "7")]

    vardiya = []

    for row in matching_rows:
        row_vardiya = []
        for col in range(13, 34):
            row_vardiya.append(worksheet.cell(row=row, column=col).value)
        vardiya.append(row_vardiya)

    for i in range(len(vardiya)):
        for j in range(len(vardiya[i])):
            if vardiya[i][j] is None:
                vardiya[i][j] = 0

    mesai_list = []
    for i in range(len(df)):
        for j in range(len(df.columns)):
            if df.iloc[i, j] == "Var":
                mesai_list.append((i * 3) + j - 1)

    for k in range(len(vardiya)):
        for i in mesai_list:
            vardiya[k][i - 1] = int(vardiya[k][i]) + int(vardiya[k][i - 1])
            vardiya[k][i] = 0

    for row, lists in zip(matching_rows, vardiya):
        for col, element in zip(range(13, 34), lists):
            worksheet.cell(row=row, column=col).value = element

    workbook.save(new_input_excel_path)

import pandas as pd
import datetime

try:
    file = pd.read_excel('Data/KW47_V00.xlsx')
    pipes = pd.read_excel('Data/Cihazlar - Borular.xlsx')
except FileNotFoundError:
    print("File not found!")
    exit(1)

indices = file.iloc[:, [0, 7, 8, 11]].reset_index()
days = file.iloc[:, 12: 33].reset_index()
combined = pd.concat([indices, days], axis=1).iloc[2:, :]

combined = combined[combined.iloc[:, 1].notna()]
combined = combined[combined.iloc[:, 2].notna()]
combined.iloc[:, 2] = combined.iloc[:, 2].astype(str)
combined = combined[combined.iloc[:, 2].apply(str.isnumeric)]

combined = combined[combined.iloc[:, 6].apply(lambda x: (type(x) != datetime.datetime) and (type(x) != str))]
combined = combined[combined.iloc[:, 3].notna()]

combined.drop('index', axis=1, inplace=True)
combined.reset_index(drop=True, inplace=True)

initial_indices = ["Hat", "Cihaz TTNr", "Cihaz Aile", "Toplam Adet"]

week_days = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
shifts = ["1", "2", "3"]

final_indices = [" ".join([i, j]) for i in week_days for j in shifts]
initial_indices.extend(final_indices)

combined = combined.set_axis(initial_indices, axis=1)
combined.set_index("Hat")

combined["Cihaz TTNr"] = combined["Cihaz TTNr"].astype(str)
pipes["Cihaz"] = pipes["Cihaz"].astype(str)

combined = combined.merge(pipes, left_on="Cihaz TTNr", right_on="Cihaz", how="inner")
combined.drop("Cihaz", axis=1, inplace=True)
combined.insert(3, 'Boru', combined.pop('Boru'))

for i in range(combined.shape[0]):
    combined.loc[i, "Pazartesi 1": "Pazar 3"] *= combined.loc[i, "Miktar"]

combined.drop("Miktar", axis=1, inplace=True)
combined.drop("Toplam Adet", axis=1, inplace=True)
combined.set_index("Hat", inplace=True)

try:
    combined.to_excel("Data/initial.xlsx")
    print("Success!")
except PermissionError:
    print("Failed!")

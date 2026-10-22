"""Optional desktop front end for the generic CSV-to-Excel demo."""
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from helper import transform


def main():
    root = tk.Tk()
    root.title("Production Planning Demo")
    root.geometry("570x200")
    folder = tk.StringVar(value=str(Path(__file__).resolve().parents[1] / "examples"))
    ttk.Label(root, text="Folder containing schedule.csv and bill_of_materials.csv").pack(pady=16)
    ttk.Entry(root, textvariable=folder, width=68).pack(padx=16)

    def choose():
        selected = filedialog.askdirectory(parent=root)
        if selected:
            folder.set(selected)

    def export():
        destination = filedialog.asksaveasfilename(
            parent=root, defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")]
        )
        if not destination:
            return
        try:
            _, _, missing = transform(folder.get(), destination)
        except (OSError, ValueError) as error:
            messagebox.showerror("Export failed", str(error), parent=root)
        else:
            messagebox.showinfo("Export complete", f"Unmapped schedule rows: {len(missing)}", parent=root)

    ttk.Button(root, text="Choose input folder", command=choose).pack(pady=8)
    ttk.Button(root, text="Export workbook", command=export).pack()
    root.mainloop()


if __name__ == "__main__":
    main()

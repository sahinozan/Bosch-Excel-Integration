from transformer_ui import show_error
import sys


# TODO: Refactor this function (elegant solution)
def path_validation(paths: dict):
    if "Source1" in paths.keys() and "Source2" in paths.keys() and "Output" in paths.keys():
        if paths["Source1"] == "" and paths["Source2"] == "":
            show_error("You have not selected both of the Excel files!")
            sys.exit(0)
        elif paths["Source1"] == "" and paths["Output"] == "":
            show_error("You have not selected the first Excel file and the output destination!")
            sys.exit(0)
        elif paths["Source2"] == "" and paths["Output"] == "":
            show_error("You have not selected the second Excel file and the output destination!")
            sys.exit(0)
        if paths["Source1"] == "":
            show_error("You have not selected the first Excel file!")
            sys.exit(0)
        elif paths["Source2"] == "":
            show_error("You have not selected the second Excel file!")
            sys.exit(0)
        elif paths["Output"] == "":
            show_error("You have not selected the output destination!")
            sys.exit(0)
    elif "Source1" in paths.keys() and "Output" in paths.keys() and "Source2" not in paths.keys():
        if paths["Source1"] == "" and paths["Output"] == "":
            show_error("You have not selected both of the Excel files and the output destination!")
            sys.exit(0)
        elif paths["Source1"] == "":
            show_error("You have not selected both of the Excel files!")
            sys.exit(0)
        elif paths["Output"] == "":
            show_error("You have not selected the second Excel file and the output destination!")
            sys.exit(0)
        show_error("You have not selected the second Excel file!")
        sys.exit(0)
    elif "Source2" in paths.keys() and "Output" in paths.keys() and "Source1" not in paths.keys():
        if paths["Source2"] == "":
            show_error("You have not selected both of the Excel files!")
            sys.exit(0)
        elif paths["Output"] == "":
            show_error("You have not selected the first Excel file and the output destination!")
            sys.exit(0)
        show_error("You have not selected the first Excel file!")
        sys.exit(0)
    elif "Source1" in paths.keys() and "Source2" not in paths.keys() and "Output" not in paths.keys():
        if paths["Source1"] == "":
            show_error("You have not selected both of the Excel files and the output destination!")
            sys.exit(0)
        show_error("You have not selected the second Excel file and the output destination!")
        sys.exit(0)
    elif "Source2" in paths.keys() and "Source1" not in paths.keys() and "Output" not in paths.keys():
        if paths["Source2"] == "":
            show_error("You have not selected both of the Excel files and the output destination!")
            sys.exit(0)
        show_error("You have not selected the first Excel file and the output destination!")
        sys.exit(0)
    elif "Source1" not in paths.keys() and "Source2" not in paths.keys() and "Output" in paths.keys():
        if paths["Output"] == "":
            show_error("You have not selected both of the Excel files and the output destination!")
            sys.exit(0)
        show_error("You have not selected the first Excel file!")
        sys.exit(0)
    elif "Source2" in paths.keys() and "Output" not in paths.keys() and "Source1" in paths.keys():
        if paths["Source1"] == "" and paths["Source2"] == "":
            show_error("You have not selected both of the Excel files and the output destination!")
            sys.exit(0)
        elif paths["Source1"] == "":
            show_error("You have not selected the first Excel file and the output destination!")
            sys.exit(0)
        elif paths["Source2"] == "":
            show_error("You have not selected the second Excel file and the output destination!")
            sys.exit(0)
        show_error("You have not selected the output destination!")
        sys.exit(0)
    elif "Source2" not in paths.keys() and "Source1" not in paths.keys() and "Output" not in paths.keys():
        show_error("You have not selected both of the Excel files and the output destination!")
        sys.exit(0)

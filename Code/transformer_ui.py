from tkinter import *
from tkinter import filedialog, messagebox
import os

# initial directory for the file explorer
current_directory = os.path.dirname(os.getcwd() + f"{os.sep}Data{os.sep}Source{os.sep}")
first_file_name, second_file_name, directory_name = "", "", ""


# Reading the file from console will be replaced with a better solution later
# TODO: Read file names without using console for the standalone executable
# TODO: Add a progress bar for the conversion process
# TODO: Polish the UI and make it more appealing with Bosch colors


def create_labels():
    input1 = Label(root, font=30, borderwidth=2, anchor="center", relief="ridge", wraplength=550)
    input2 = Label(root, font=30, borderwidth=2, anchor="center", relief="ridge", wraplength=550)
    output = Label(root, font=30, borderwidth=2, anchor="center", relief="ridge", wraplength=550)
    progress = Label(root, font=30, borderwidth=2, anchor="center", relief="ridge",
                     text="Progress bar is not implemented yet")
    return input1, input2, output, progress


# browse the input file (opens a file dialog to check the file name)
def browse_first_input_file():
    global first_file_name
    first_file_name = filedialog.askopenfilename(initialdir=current_directory,
                                                 title="Select a File",
                                                 filetypes=(("Excel Files", "*.xlsx"),
                                                            ("Excel Macro Files", "*.xlsm"),))
    print(f"Source1={first_file_name}")
    if first_file_name != "":
        input1_ent.config(text=first_file_name)


def browse_second_input_file():
    global second_file_name
    second_file_name = filedialog.askopenfilename(initialdir=current_directory,
                                                  title="Select a File",
                                                  filetypes=(("Excel Files", "*.xlsx"),
                                                             ("Excel Macro Files", "*.xlsm"),))
    print(f"Source2={second_file_name}")
    if second_file_name != "":
        input2_ent.config(text=second_file_name)


# browse the output destination (opens a file dialog to check the directory name)
def browse_output_destination():
    global directory_name
    directory_name = filedialog.askdirectory(initialdir=current_directory,
                                             title="Select a Directory")
    print(f"Output={directory_name}")
    if directory_name != "":
        output_ent.config(text=directory_name)


def show_error(error_message):
    root.withdraw()
    messagebox.showerror(title="Error", message=error_message, icon="error")
    root.deiconify()


# create the user interface
def create_ui():
    # Set window title
    root.title('Excel File Explorer')

    w = 750  # width for the Tk root
    h = 300  # height for the Tk root

    # get screen width and height
    ws = root.winfo_screenwidth()  # width of the screen
    hs = root.winfo_screenheight()  # height of the screen

    # calculate x and y coordinates for the Tk root window
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)

    # set the dimensions of the screen where the window will be displayed
    root.geometry('%dx%d+%d+%d' % (w, h, x, y))

    # center the buttons and window
    root.grid_columnconfigure(0, weight=0)
    root.grid_columnconfigure(1, weight=2)

    # make the windows expandable
    root.grid_rowconfigure(0, weight=1)
    root.grid_rowconfigure(1, weight=1)
    root.grid_rowconfigure(2, weight=1)
    root.grid_rowconfigure(3, weight=1)
    root.grid_rowconfigure(4, weight=1)

    # make the window non-resizable
    root.resizable(False, False)

    button_first_source = Button(root,
                                 text="Next Week",
                                 command=browse_first_input_file)

    button_second_source = Button(root,
                                  text="Past Week",
                                  command=browse_second_input_file)

    button_output = Button(root,
                           text="Output Destination",
                           command=browse_output_destination)

    button_exit = Button(root,
                         text="Start Transformation",
                         command=root.destroy)

    # Place the buttons and text boxes in the window
    button_first_source.grid(row=0, column=0, sticky="nsew", pady=1, padx=2)
    input1_ent.grid(row=0, column=1, sticky="nsew", pady=4, padx=5)
    button_second_source.grid(row=1, column=0, sticky="nsew", pady=1, padx=2)
    input2_ent.grid(row=1, column=1, sticky="nsew", pady=4, padx=5)
    button_output.grid(row=2, column=0, sticky="nsew", pady=1, padx=2)
    output_ent.grid(row=2, column=1, sticky="nsew", pady=4, padx=5)
    button_exit.grid(row=3, columnspan=2, sticky="nsew", pady=3, padx=2)
    progress_bar.grid(row=4, columnspan=2, sticky="nsew", pady=5, padx=5)

    # Let the window wait for any events
    root.mainloop()


if __name__ == "__main__":
    # create the window
    root = Tk()

    # create the labels
    input1_ent, input2_ent, output_ent, progress_bar = create_labels()

    # create the user interface
    create_ui()

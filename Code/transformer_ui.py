from tkinter import *
from tkinter import filedialog
import os

current_directory = os.path.dirname(os.getcwd() + f"{os.sep}Data{os.sep}Source{os.sep}")
file_name = ""
directory_name = ""


def browse_input_file():
    global file_name
    file_name = filedialog.askopenfilename(initialdir=current_directory,
                                           title="Select a File",
                                           filetypes=(("Excel Files", "*.xlsx"),
                                                      ("Excel Macro Files", "*.xlsm"),))
    print(f"Source={file_name}")


def browse_output_destination():
    global directory_name
    directory_name = filedialog.askdirectory(initialdir=current_directory,
                                             title="Select a Directory")
    print(f"Output={directory_name}")


def create_ui():

    # Create the root window
    global root
    root = Tk()

    # Set window title
    root.title('Excel File Explorer')

    w = 300  # width for the Tk root
    h = 200  # height for the Tk root

    # get screen width and height
    ws = root.winfo_screenwidth()  # width of the screen
    hs = root.winfo_screenheight()  # height of the screen

    # calculate x and y coordinates for the Tk root window
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)

    # set the dimensions of the screen where the window will be displayed
    root.geometry('%dx%d+%d+%d' % (w, h, x, y))

    # center the buttons and window
    root.grid_columnconfigure(10, weight=1)
    root.grid_rowconfigure(10, weight=1)

    # make the window non-resizable
    root.resizable(0, 0)

    button_source = Button(root,
                           text="Select Excel File",
                           command=browse_input_file)

    button_output = Button(root,
                           text="Select Output Destination",
                           command=browse_output_destination)

    button_exit = Button(root,
                         text="Start Transformation",
                         command=root.destroy)

    button_source.pack(side=TOP, anchor="center",
                       expand=True, fill="both", padx=10, pady=5)

    button_output.pack(side=TOP, anchor="center",
                       expand=True, fill="both", padx=10, pady=5)

    button_exit.pack(side=BOTTOM, anchor="center",
                     expand=True, fill="both", padx=10, pady=5)

    # Let the window wait for any events
    root.mainloop()


if __name__ == "__main__":
    create_ui()

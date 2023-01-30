from tkinter import *  # type: ignore
from tkinter import filedialog
from helper import resource_path
import os

# Reading the file from console will be replaced with a better solution later
# TODO: Read file names without using console for the standalone executable
# TODO: Add a text bar for the selected input file directory in the UI
# TODO: Add a text bar for the selected destination directory in the UI
# TODO: Add a progress bar for the conversion process
# TODO: Polish the UI and make it more appealing with Bosch colors

# initial directory for the file explorer
current_directory = os.path.join(os.path.join(os.environ['HOME']), 'Desktop')


# browse the input file (opens a file dialog to check the file name)
def browse_input_file():
    global file_name
    file_name = filedialog.askopenfilename(initialdir=current_directory,
                                           title="Select a File",
                                           filetypes=(("Excel Files", "*.xlsx"),
                                                      ("Excel Macro Files", "*.xlsm"),))
    with open(resource_path("paths.txt"), "w") as file:
        file.write(file_name + ";")


# browse the output destination (opens a file dialog to check the directory name)
def browse_output_destination():
    global directory_name
    directory_name = filedialog.askdirectory(initialdir=current_directory,
                                             title="Select a Directory")
    with open(resource_path("paths.txt"), "a") as file:
        file.write(directory_name)


# create the user interface
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
    root.resizable(0, 0)  # type: ignore

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

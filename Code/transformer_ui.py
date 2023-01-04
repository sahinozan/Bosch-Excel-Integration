from tkinter import *  # type: ignore
from tkinter import filedialog
import os

current_directory = os.path.dirname(os.getcwd() + "/Data/Source/")
filename = ""


def browseFiles():
    global filename
    filename = filedialog.askopenfilename(initialdir=current_directory,
                                          title="Select a File",
                                          filetypes=(("Excel files",
                                                      "*.xlsx"),))
    print(filename)
    root.destroy()


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

    button_explore = Button(root,
                            text="Browse Files",
                            command=browseFiles)

    button_explore.pack(side=TOP, anchor="center",
                        expand=True, fill="both", padx=10, pady=10)

    # Let the window wait for any events
    root.mainloop()


if __name__ == "__main__":
    create_ui()

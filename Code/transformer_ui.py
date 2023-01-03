from tkinter import *  # type: ignore
from tkinter import filedialog
import os

current_directory = os.path.dirname(os.getcwd() + "/Data/Source/")
filename = ""

def browseFiles():
    global filename
    filename = filedialog.askopenfilename(initialdir = current_directory,
                                          title = "Select a File",
                                          filetypes = (("Excel files",
                                                        "*.xlsx"),))
    print(filename)
    window.destroy()


def create_ui():
                                                                                              
    # Create the root window
    global window
    window = Tk()
    
    # Set window title
    window.title('Excel File Explorer')
    
    # Set window size
    window.geometry("300x200")
    
    # center the buttons and window
    window.grid_columnconfigure(10, weight=1)
    window.grid_rowconfigure(10, weight=1)

    # make the window non-resizable
    window.resizable(0, 0)  # type: ignore
        
    button_explore = Button(window,
                            text = "Browse Files",
                            command = browseFiles)

    button_explore.pack(side=TOP, anchor="center", 
                        expand=True, fill="both", padx=20, pady=20)
    
    # Let the window wait for any events
    window.mainloop()

if __name__ == "__main__":
    create_ui()

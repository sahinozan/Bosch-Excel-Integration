from tkinter import *
from tkinter import filedialog, messagebox
import os
import sys


def show_error(message):
    message_window = Tk()
    message_window.withdraw()
    messagebox.showerror("Error", message, icon="error")
    message_window.destroy()
    sys.exit(0)


class TransformerUI:
    def __init__(self, root_window):
        self.current_directory = os.path.dirname(os.getcwd() + f"{os.sep}Data{os.sep}Source{os.sep}")
        self.first_file_name, self.second_file_name, self.directory_name = "", "", ""
        self.close = False
        self.progress_bar = None
        self.root = root_window
        self.window_configuration()
        self.grid_configuration()
        self.input1, self.input2, self.output, self.progress = self.create_labels()
        self.input1_button, self.input2_button, self.output_button, self.transform_button = self.create_buttons()
        self.place_components()

    def on_close(self):
        self.close = messagebox.askokcancel("Close", "Would you like to close the program?",
                                            icon="warning", parent=root)
        if self.close:
            self.root.destroy()
            sys.exit(0)

    @staticmethod
    def default_font():
        if sys.platform == "win32":
            return "Arial", int(10)
        elif sys.platform == "darwin":
            return "Courier", int(13)

    def create_buttons(self):
        input1_button = Button(self.root, text="Next Week", font=(self.default_font()),
                               command=self.browse_first_input_file)
        input2_button = Button(self.root, text="Past Week", font=(self.default_font()),
                               command=self.browse_second_input_file)
        output_button = Button(self.root, text="Output Destination", font=(self.default_font()),
                               command=self.browse_output_directory)
        transform_button = Button(self.root, text="Transform", font=(self.default_font()),
                                  command=root.destroy)
        return input1_button, input2_button, output_button, transform_button

    def create_labels(self):
        input1 = Label(self.root, font=(self.default_font()), borderwidth=2, anchor="center",
                       relief="ridge", wraplength=550)
        input2 = Label(self.root, font=(self.default_font()), borderwidth=2, anchor="center",
                       relief="ridge", wraplength=550)
        output = Label(self.root, font=(self.default_font()), borderwidth=2, anchor="center",
                       relief="ridge", wraplength=550)
        progress = Label(self.root, font=(self.default_font()), borderwidth=2, anchor="center",
                         relief="ridge", text="Progress bar is not implemented yet")
        return input1, input2, output, progress

    def place_components(self):
        self.input1_button.grid(row=0, column=0, padx=2, pady=1, sticky="nsew")
        self.input1.grid(row=0, column=1, padx=5, pady=4, sticky="nsew")

        self.input2_button.grid(row=1, column=0, padx=2, pady=1, sticky="nsew")
        self.input2.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")

        self.output_button.grid(row=2, column=0, padx=2, pady=1, sticky="nsew")
        self.output.grid(row=2, column=1, padx=5, pady=4, sticky="nsew")

        self.transform_button.grid(row=3, columnspan=2, padx=3, pady=2, sticky="nsew")
        self.progress.grid(row=4, columnspan=2, padx=5, pady=5, sticky="nsew")

    def grid_configuration(self):
        # center the buttons and window
        self.root.grid_columnconfigure(0, weight=0)
        self.root.grid_columnconfigure(1, weight=2)

        # make the windows expandable
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_rowconfigure(2, weight=1)
        self.root.grid_rowconfigure(3, weight=1)
        self.root.grid_rowconfigure(4, weight=1)

    def window_configuration(self):
        self.root.title("Excel File Explorer")

        # make the window non-resizable
        self.root.resizable(False, False)

        # make the window close when the user clicks the X button
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # root.overrideredirect(True)

        w = 750  # width for the Tk root
        h = 300  # height for the Tk root

        # get screen width and height
        ws = self.root.winfo_screenwidth()  # width of the screen
        hs = self.root.winfo_screenheight()  # height of the screen

        # calculate x and y coordinates for the Tk root window
        x = (ws / 2) - (w / 2)
        y = (hs / 2) - (h / 2)

        # set the dimensions of the screen where the window will be displayed
        self.root.geometry('%dx%d+%d+%d' % (w, h, x, y))

    def browse_first_input_file(self):
        self.first_file_name = filedialog.askopenfilename(initialdir=self.current_directory,
                                                          title="Select a File",
                                                          filetypes=(("Excel Files", "*.xlsx"),
                                                                     ("Excel Macro Files", "*.xlsm"),))
        print(f"Source1={self.first_file_name}")
        if self.first_file_name != "":
            self.input1.config(text=self.first_file_name)

    def browse_second_input_file(self):
        self.second_file_name = filedialog.askopenfilename(initialdir=self.current_directory,
                                                           title="Select a File",
                                                           filetypes=(("Excel Files", "*.xlsx"),
                                                                      ("Excel Macro Files", "*.xlsm"),))
        print(f"Source2={self.second_file_name}")
        if self.second_file_name != "":
            self.input2.config(text=self.second_file_name)

    def browse_output_directory(self):
        self.directory_name = filedialog.askdirectory(initialdir=self.current_directory,
                                                      title="Select a Directory")
        print(f"Output={self.directory_name}")
        if self.directory_name != "":
            self.output.config(text=self.directory_name)


if __name__ == "__main__":
    root = Tk()
    app = TransformerUI(root)
    root.mainloop()

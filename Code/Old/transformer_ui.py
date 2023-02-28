from tkinter import *
from tkinter import filedialog, messagebox
import os
import sys


class TransformerUI:
    def __init__(self, root_window) -> None:
        """
        Initializes the UI.

        Args:
            root_window: The root window of the application.

        Attributes:
            current_directory: The current directory of the application. (e.g. /Users/user/Documents/)
            first_file_name: The name of the first Excel file which contains the next week's production plan.
            second_file_name: The name of the second Excel file which contains the current week's production plan.
            directory_name: The name of the output directory where the transformed Excel file will be saved.
            close: A boolean value that determines whether the program will be closed.
            progress_bar: The progress bar of the application.
            root: The root window of the application.
        """
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

    def on_close(self) -> None:
        """
        Asks the user whether they want to close the program.
        """
        self.close = messagebox.askokcancel("Close", "Would you like to close the program?",
                                            icon="warning", parent=self.root)

        if self.close:
            self.root.destroy()
            sys.exit(0)

    @staticmethod
    def show_error(message) -> None:
        """
        Shows an error message.

        Args:
            message: The error message that will be displayed. (e.g. "You have not selected the first Excel file!")
        """
        message_window = Tk()
        message_window.withdraw()
        messagebox.showerror("Error", message, icon="error")
        message_window.destroy()
        sys.exit(0)

    @staticmethod
    def default_font() -> tuple[str, int]:
        """
        Checks whether the operating system is Windows or macOS and selects the default font and font size accordingly.

        Returns:
            A tuple with the default font and font size.
        """
        if sys.platform == "win32":
            return "Arial", int(10)
        elif sys.platform == "darwin":
            return "Courier", int(13)

    def create_buttons(self) -> tuple[Button, Button, Button, Button]:
        """
        Creates the buttons in the UI. Input1 and input2 are the buttons that are used to select the Excel files for
        the next week's and current week's production plan. Output is the button that is used to select the output
        directory for the transformed Excel file. Transform is the button that is used to start the transformation.

        Returns:
            A tuple with the next week, current week, output destination, and transform buttons.
        """
        input1_button = Button(self.root, text="Next Week", font=(self.default_font()),
                               command=self.browse_first_input_file)
        input2_button = Button(self.root, text="Current Week", font=(self.default_font()),
                               command=self.browse_second_input_file)
        output_button = Button(self.root, text="Output Destination", font=(self.default_font()),
                               command=self.browse_output_directory)
        transform_button = Button(self.root, text="Transform", font=(self.default_font()),
                                  command=self.root.destroy)
        return input1_button, input2_button, output_button, transform_button

    def create_labels(self) -> tuple[Label, Label, Label, Label]:
        """
        Creates the labels in the UI. Input1 and input2 are the text boxes where the user can see the selected file
        path. Output is the text box where the user can see the selected output directory. Progress is the text bar
        that is used to show the progress of the transformation.

        Returns:
            A tuple with the next week, current week, output destination, and progress labels.
        """
        input1 = Label(self.root, font=(self.default_font()), borderwidth=2, anchor="center",
                       relief="ridge", wraplength=550)
        input2 = Label(self.root, font=(self.default_font()), borderwidth=2, anchor="center",
                       relief="ridge", wraplength=550)
        output = Label(self.root, font=(self.default_font()), borderwidth=2, anchor="center",
                       relief="ridge", wraplength=550)
        progress = Label(self.root, font=(self.default_font()), borderwidth=2, anchor="center",
                         relief="ridge", text="Progress bar is not implemented yet")
        return input1, input2, output, progress

    def place_components(self) -> None:
        """
        Places the components such as buttons and labels in the UI.
        """
        self.input1_button.grid(row=0, column=0, padx=2, pady=1, sticky="nsew")
        self.input1.grid(row=0, column=1, padx=5, pady=4, sticky="nsew")

        self.input2_button.grid(row=1, column=0, padx=2, pady=1, sticky="nsew")
        self.input2.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")

        self.output_button.grid(row=2, column=0, padx=2, pady=1, sticky="nsew")
        self.output.grid(row=2, column=1, padx=5, pady=4, sticky="nsew")

        self.transform_button.grid(row=3, columnspan=2, padx=3, pady=2, sticky="nsew")
        self.progress.grid(row=4, columnspan=2, padx=5, pady=5, sticky="nsew")

    def grid_configuration(self) -> None:
        """
        Configures the row and column weights of the grid. The grid will be placed in the window to align the buttons
        and the labels of the UI.
        """
        # center the buttons and window
        self.root.grid_columnconfigure(0, weight=0)
        self.root.grid_columnconfigure(1, weight=2)

        # make the windows expandable
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_rowconfigure(2, weight=1)
        self.root.grid_rowconfigure(3, weight=1)
        self.root.grid_rowconfigure(4, weight=1)

    def window_configuration(self) -> None:
        """
        Sets the title of the window, makes the window non-resizable, and makes the window close when the user clicks.
        Also sets the dimensions of the window and places it in the center of the screen.
        """
        self.root.title("Excel Converter")

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

    def browse_first_input_file(self) -> None:
        """
        Opens a file dialog box to select the next week's production plan Excel file. The file dialog box will only
        show Excel files.
        """
        self.first_file_name = filedialog.askopenfilename(initialdir=self.current_directory,
                                                          title="Select a File (Next Week)",
                                                          filetypes=(("Excel Files", "*.xlsx"),
                                                                     ("Excel Macro Files", "*.xlsm"),))
        print(f"Source1={self.first_file_name}")
        if self.first_file_name != "":
            self.input1.config(text=self.first_file_name)

    def browse_second_input_file(self) -> None:
        """
        Opens a file dialog box to select the current week's production plan Excel file. The file dialog box will
        only show Excel files.
        """
        self.second_file_name = filedialog.askopenfilename(initialdir=self.current_directory,
                                                           title="Select a File (Current Week)",
                                                           filetypes=(("Excel Files", "*.xlsx"),
                                                                      ("Excel Macro Files", "*.xlsm"),))
        print(f"Source2={self.second_file_name}")
        if self.second_file_name != "":
            self.input2.config(text=self.second_file_name)

    def browse_output_directory(self) -> None:
        """
        Opens a directory dialog box to select the output directory where the transformed Excel file will be saved.
        """
        self.directory_name = filedialog.askdirectory(initialdir=self.current_directory,
                                                      title="Select a Directory")
        print(f"Output={self.directory_name}")
        if self.directory_name != "":
            self.output.config(text=self.directory_name)


if __name__ == "__main__":
    root = Tk()
    app = TransformerUI(root)
    root.mainloop()

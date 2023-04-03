# Author: Ozan Şahin

import os
import sys
from tkinter import messagebox, filedialog, Tk
from customtkinter import CTk, CTkFrame, CTkLabel, CTkButton, CTkFont, CTkOptionMenu, CTkEntry, set_appearance_mode, \
    set_default_color_theme


class App(CTk):
    set_appearance_mode("Dark")
    set_default_color_theme("dark-blue")

    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)

        self.current_directory = os.path.dirname(os.getcwd() + f"{os.sep}Data{os.sep}Source{os.sep}")
        self.first_file_name, self.second_file_name, self.directory_name = "", "", ""

        self.shift_window = None
        self.close = None
        self.window_configuration()
        self.grid_configuration()

        # create sidebar frame with widgets
        self.sidebar_frame = CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=3, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(6, weight=1)
        self.logo_label = CTkLabel(self.sidebar_frame, text="MOE32\nProduction\nPlanner",
                                   font=CTkFont(size=17, weight="bold"), anchor="w")
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.input_file_button, self.output_destination_button, self.transform_button = self.create_buttons()
        self.input_file_path, self.output_destination_path = self.create_labels()
        self.place_components()

        self.appearance_mode_label = CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_option_menu = CTkOptionMenu(self.sidebar_frame,
                                                         values=["Dark", "Light", "System"],
                                                         command=self.change_appearance_mode_event,
                                                         font=("Arial", 12),
                                                         anchor="w")
        self.appearance_mode_option_menu.grid(row=8, column=0, padx=20, pady=(10, 20))
        self.create_labels()

        self.button_validation()

    def button_validation(self) -> None:
        """
        Sets the default values of the widgets and controls the button activation.
        """
        self.input_file_path.configure(justify="center")
        self.output_destination_path.configure(justify="center")
        self.appearance_mode_option_menu.set("Dark")

        # check the current values and activate the buttons accordingly
        if self.input_file_path.get() == "" or self.input_file_path.get() == "Input File Path":
            self.transform_button.configure(state="disabled")
        else:
            self.transform_button.configure(state="normal")

    def on_close(self) -> None:
        """
        Asks the user whether they want to close the program.
        """
        self.close = messagebox.askokcancel("Close", "Would you like to close the program?",
                                            icon="warning", parent=self)

        if self.close:
            self.destroy()
            sys.exit(0)

    @staticmethod
    def show_error(message: str, system_exit=True) -> None:
        """
        Shows an error message.

        Args:
            message: The error message that will be displayed. (e.g. "You have not selected the first Excel file!")
            system_exit: Whether the program should be closed after the error message has been displayed.
        """
        message_window = CTk()
        message_window.withdraw()
        messagebox.showerror("Error", message, icon="error")
        message_window.destroy()
        if system_exit:
            sys.exit(0)

    def show_info(self, message: str, system_exit=True) -> None:
        """
        Shows an information message.

        Args:
            message: The error message that will be displayed. (e.g. "You have not selected the first Excel file!")
            system_exit: Whether the program should be closed after the error message has been displayed.
        """
        message_window = Tk()

        message_window.withdraw()
        messagebox.showinfo(title="Information", message=message, icon="info", parent=self)
        message_window.destroy()
        if system_exit:
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

    def create_buttons(self) -> tuple[CTkButton, CTkButton, CTkButton]:
        """
        Creates the buttons in the UI. Input1 and input2 are the buttons that are used to select the Excel files for
        the next week's and current week's production plan. Output is the button that is used to select the output
        directory for the transformed Excel file. Transform is the button that is used to start the transformation.

        Returns:
            A tuple with the next week, current week, output destination, and transform buttons.
        """
        # second_rule_button = CTkButton(self.sidebar_frame, command=self.get_shifts,
        #                                text="Shift Control Rule", font=("Arial", 12))
        input_file_button = CTkButton(self.sidebar_frame, command=self.browse_input_file,
                                      text="Next Week", font=("Arial", 12))
        output_destination_button = CTkButton(self.sidebar_frame, command=self.browse_output_directory,
                                              text="Output Destination", font=("Arial", 12))
        transform_button = CTkButton(self.sidebar_frame, command=self.destroy,
                                     text="Transform", font=("Arial", 12))
        return input_file_button, output_destination_button, transform_button

    def create_labels(self) -> tuple[CTkEntry, CTkEntry]:
        """
        Creates the labels in the UI. Input1 and input2 are the text boxes where the user can see the selected file
        path. Output is the text box where the user can see the selected output directory. Progress is the text bar
        that is used to show the progress of the transformation.

        Returns:
            A tuple with the next week, current week, output destination, and progress labels.
        """
        input_file_path = CTkEntry(self, placeholder_text="Source Excel Path")
        input_file_path.configure(state="disabled")
        input_file_path.xview_scroll(1, "units")
        output_destination_path = CTkEntry(self, placeholder_text="Output Destination")
        output_destination_path.configure(state="disabled")
        return input_file_path, output_destination_path

    def place_components(self) -> None:
        """
        Places the components such as buttons and labels in the UI.
        """
        self.input_file_button.grid(row=1, column=0, padx=20, pady=10)
        self.output_destination_button.grid(row=2, column=0, padx=20, pady=10)
        self.transform_button.grid(row=3, column=0, padx=20, pady=10)
        self.input_file_path.grid(row=0, column=1, padx=(20, 20), pady=(50, 0), sticky="nsew")
        self.output_destination_path.grid(row=2, column=1, padx=(20, 20), pady=(0, 50), sticky="nsew")

    def grid_configuration(self) -> None:
        """
        Configures the row and column weights of the grid. The grid will be placed in the window to align the buttons
        and the labels of the UI.
        """
        # center the buttons and window
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=0)
        self.grid_columnconfigure(3, weight=0)

        # make the windows expandable
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=0)

    def window_configuration(self) -> None:
        """
        Sets the title of the window, makes the window non-resizable, and makes the window close when the user clicks.
        Also sets the dimensions of the window and places it in the center of the screen.
        """
        self.title("MOE32 Production Planner")

        # make the window non-resizable
        self.resizable(False, False)

        # make the window close when the user clicks the X button
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # set the dimensions of the window
        w = 750  # width for the Tk root
        h = 330  # height for the Tk root

        # get screen width and height
        ws = self.winfo_screenwidth()  # width of the screen
        hs = self.winfo_screenheight()  # height of the screen

        # calculate x and y coordinates for the Tk root window
        x = (ws / 2) - (w / 2)
        y = (hs / 2) - (h / 2)

        # set the dimensions of the screen where the window will be displayed
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))

    def browse_input_file(self) -> None:
        """
        Opens a file dialog box to select the next week's production plan Excel file. The file dialog box will only
        show Excel files.
        """
        self.first_file_name = filedialog.askopenfilename(initialdir=self.current_directory,
                                                          title="Select a File",
                                                          filetypes=(("Excel Files", "*.xlsx"),
                                                                     ("Excel Macro Files", "*.xlsm"),))
        print(f"Source={self.first_file_name}")
        self.input_file_path.configure(state="normal")
        if self.input_file_path.get() != "":
            self.input_file_path.delete(0, "end")
        if self.first_file_name != "":
            self.input_file_path.insert(index=0, string=self.first_file_name)
        else:
            self.input_file_path.insert(index=0, string="Next Week")
        self.input_file_path.configure(state="disabled")
        self.button_validation()

    def browse_output_directory(self) -> None:
        """
        Opens a directory dialog box to select the output directory where the transformed Excel file will be saved.
        """
        self.directory_name = filedialog.askdirectory(initialdir=self.current_directory,
                                                      title="Select a Directory")
        print(f"Output={self.directory_name}")
        self.output_destination_path.configure(state="normal")
        if self.output_destination_path.get() != "":
            self.output_destination_path.delete(0, "end")
        if self.directory_name != "":
            self.output_destination_path.insert(index=0, string=self.directory_name)
        else:
            self.output_destination_path.insert(index=0, string="Output Destination")
        self.output_destination_path.configure(state="disabled")
        self.button_validation()

    @staticmethod
    def change_appearance_mode_event(new_appearance_mode: str) -> None:
        """
        Changes the appearance mode of the UI.

        Args:
            new_appearance_mode: The new appearance mode of the UI. (light or dark)
        """
        set_appearance_mode(new_appearance_mode)


if __name__ == "__main__":
    app = App()
    app.mainloop()

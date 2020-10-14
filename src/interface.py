import csv
import datetime
import glob
import os
import shutil  # Delete folder
import threading
import time
import tkinter as tk
import webbrowser
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

import requests

from src import receiver, previous_semesters, previous_data, room_capacity


class UserInterface(Frame):
    # It is better to define values like the following ones as constants (uppercase) in a single place (like here)
    GOOGLE_FORM_URL = 'https://goo.gl/forms/wNkzjymOQ7wiNavf1'
    INSTRUCTIONS_URL = 'https://docs.google.com/document/d/1htRsKmxDX33yawrYqeHkCLWlEL-juRjeM-if8N4f2yo/edit?usp=sharing'
    LATEST_RELEASE_URL = 'https://github.com/igorneaga/schedule/releases/latest'

    def __init__(self, master, current_path):
        super().__init__(master)
        self.grid()

        # Assets
        self.cwd = current_path
        self.path = self.cwd

        self.BackImage = tk.PhotoImage(file=f'{self.cwd}\\assets\\back_icon_45x45.png')
        self.OutOrderImage = tk.PhotoImage(file=f'{self.cwd}\\assets\\table_v05_default.png')
        self.InOrderImage = tk.PhotoImage(file=f'{self.cwd}\\assets\\table_v05_in_order.png')
        self.ExcelCopyFile = tk.PhotoImage(file=f'{self.cwd}\\assets\\excel_files_icon.png')
        self.ExcelMainFile = tk.PhotoImage(file=f'{self.cwd}\\assets\\master_file_icon.png')
        self.CreateMasterImage = tk.PhotoImage(file=f'{self.cwd}\\assets\\create_master.png')
        self.CreatePayrollImage = tk.PhotoImage(file=f'{self.cwd}\\assets\\create_fwm_table.png')
        self.GetPreviousImage = tk.PhotoImage(file=f'{self.cwd}\\assets\\get_prev_tables.png')
        self.ExitApplicationImage = tk.PhotoImage(file=f'{self.cwd}\\assets\\quit_button.png')
        self.UseLocalFiles = tk.PhotoImage(file=f'{self.cwd}\\assets\\use_local.png')

        # Default table characteristics
        today_date = time.strftime("%Y,%m")
        today_date_split = today_date.split(',')
        self.table_settings_type = 1
        self.table_settings_year = today_date_split[0]
        self.table_settings_semester = "Fall"
        self.table_settings_name = "Uni_Table"
        self.table_friday_include = 0

        if int(today_date_split[1]) < 5:
            self.web_semester_parameters = "Spring"
        else:
            self.web_semester_parameters = "Fall"
        self.web_department_parameters = "ACCT"
        self.web_year = today_date_split[0]
        # Holds user choice both standard and urlencode
        self.urlencode_dict_list = []

        # Information from data files
        self.file_name = None
        self.files_show_directory = []
        self.files_show_names = []
        self.files_string = None

        # GUI windows
        self.selection_window = None
        self.files_manipulation_window = None
        self.introduction = None
        self.settings_window = None
        self.creating_step_window = None
        self.notification_window = None
        self.get_example_window = None
        self.payroll_window = None
        self.cost_department_list = None

        # GUI buttons, radio buttons, insertion box and others
        self.create_table_button = None
        self.get_value = None  # Needs for radio buttons
        self.include_friday = None
        self.table_order_default = None
        self.table_order_type = None
        self.table_name_insertion_box = None

        # Cost center files & data
        self.cost_center_string = "None"
        self.file = 'cost_center.cvs'
        self.department_cost_dict = {}

        # Stores all errors found from receiver.py
        self.error_data_list = []

        # Semester & Department & Year from university website
        self.web_semesters_options = []
        self.web_department_options = []
        self.web_year_options = []

        # Stores data about room capacity
        self.room_cap_dict = room_capacity.RoomCapacity().get_capacity()

        # A label which will keep updating once user choose a data file
        self.button_text = tk.StringVar()
        self.button_text.set("File(s) Selected: ")
        self.create_files_names = Button(self.selection_window, border=0,
                                         textvariable=self.button_text, command=self.change_files_window,
                                         foreground="gray", font=("Arial", 11, "bold"))

        # User directory shortcut
        self.user_directory = "/"

        # Other
        self.payroll_selection = None
        self.listbox = None
        self.folder = None
        self.cost_box_insert = None
        self.mini_frame = None
        self.move_next_step = None
        self.payroll_year_1 = None
        self.payroll_semester_1 = None
        self.payroll_year_2 = None
        self.payroll_semester_2 = None

        # Deletes previous files
        shutil.rmtree('copy_folder', ignore_errors=True)
        shutil.rmtree('__excel_files', ignore_errors=True)
        shutil.rmtree('web_files', ignore_errors=True)

        def receive_semesters():
            try:
                p = previous_semesters.ReceiveSemesters().return_courses_semesters()
                return p
            except requests.exceptions.ConnectionError:
                messagebox.showwarning(title="Connection Error", message="Check your connection. Some functions might "
                                                                         "not work properly")

        self.param = receive_semesters()
        self.organize_semester_data()

        try:
            excel_file = glob.glob('__excel_files/*.xlsx')
            if not excel_file:
                pass
            else:
                open(excel_file[0], "r+")
        except IOError:
            messagebox.showerror("Close File", "Please close excel files to eliminate errors")
        self.introduction_window()

    def organize_semester_data(self):
        for param_len in range(len(self.param)):
            # Finds available options from scraping
            for key in self.param[param_len]:
                test_dict = dict()
                if key[0:4] == "FALL" or key[0:4] == "SPRI":
                    find_year_index = key.find("2")
                    self.web_semesters_options.append(key[0:find_year_index])
                    self.web_year_options.append(key[find_year_index:])
                    test_dict[key[0:find_year_index]] = self.param[param_len].get(key)
                    self.urlencode_dict_list.append(test_dict)

                else:
                    symbol_index = key.find("(")
                    self.web_department_options.append(key[symbol_index + 1:-1])
                    test_dict[key[symbol_index + 1:-1]] = self.param[param_len].get(key)
                    self.urlencode_dict_list.append(test_dict)

        # Insert additional shortkey departments
        self.web_department_options.append("All COB Departments")
        self.web_department_options.append("ACCT & BLAW & MACC")
        self.web_department_options.append("MRKT & IBUS")
        self.web_department_options.append("MGMT & MBA")

    def submit_ticket_form(self):
        """Opens a Google Form to collect any reports or requests"""
        webbrowser.open(self.GOOGLE_FORM_URL)

    def open_instructions_url(self):
        """Instructions on how to use this program"""
        webbrowser.open(self.INSTRUCTIONS_URL)

    def open_latest_release(self):
        webbrowser.open(self.LATEST_RELEASE_URL)

    def main_text_interface(self, button_frame, title_text, back_button_function, description_text=None,
                            x_description=18, y_title=20, remove_back=False):
        title_label = ttk.Label(button_frame,
                                text=title_text,
                                foreground="green",
                                font=('Arial', 18))
        title_label.grid(sticky='W',
                         column=0,
                         columnspan=2,
                         row=0,
                         rowspan=2,
                         padx=250,
                         pady=y_title)
        if description_text is None:
            pass
        else:
            # Description of a reason to have this window
            description_label = ttk.Label(button_frame,
                                          text=description_text,
                                          foreground="gray",
                                          font=('Arial', 12))

            description_label.grid(column=0,
                                   row=1,
                                   rowspan=2,
                                   padx=x_description,
                                   pady=75)
        if remove_back is False:
            back_button = Button(button_frame,
                                 border='0',
                                 image=self.BackImage,
                                 command=back_button_function)

            back_button.grid(sticky='WN',
                             column=0,
                             row=1,
                             rowspan=2,
                             pady=15,
                             padx=10)

    def interface_window_remover(self):
        """Removes window once a user goes to a next step or previous step."""

        if self.introduction:
            self.introduction.grid_remove()

        if self.payroll_window:
            self.payroll_window.grid_remove()
            self.cost_department_list.grid_remove()
            self.cost_dict()
            self.payroll_window = None

        if self.get_example_window:
            self.get_example_window.grid_remove()

        if self.selection_window:
            self.selection_window.grid_remove()

        if self.files_manipulation_window:
            self.files_manipulation_window.grid_remove()

        if self.settings_window:
            self.settings_window.grid_remove()

        if self.creating_step_window:
            self.creating_step_window.grid_remove()

        if self.notification_window:
            self.notification_window.grid_remove()

        self.create_files_names.place_forget()

    def select_excel_files(self):
        """Once a user selects the file - it will hold in the list."""
        self.file_name = Frame(self).filename = filedialog.askopenfilenames(initialdir=self.user_directory,
                                                                            title="Select Excel file",
                                                                            filetypes=(("excel files", "*.xlsx"),
                                                                                       ("all files", "*.*")))
        if not self.file_name:
            pass
        else:
            self.user_directory = str()
            # For display and files store
            for filesAmount in range(len(self.file_name)):
                split_user_directory = self.user_directory.split("/")
                split_user_directory = (split_user_directory[0:len(split_user_directory) - 1])
                for dir_length in range(len(split_user_directory)):
                    # Stores user directory of the previously selected file to access easily next time
                    self.user_directory += split_user_directory[dir_length] + "/"
                self.files_show_directory.append(self.file_name[filesAmount])
                self.display_excel_files()

    def display_excel_files(self):
        """Shows to the user which files has been chosen"""
        # Prepares the file names into the proper format.
        self.files_show_names = []
        for i in self.files_show_directory:
            z = 0
            for _ in i:
                z -= 1
                if i[z] == '/' or i[z] == '\\':
                    self.files_show_names.insert(0, i[z + 1:])
                    break

        if len(self.files_show_names) == 1:
            self.files_string = ("File(s) Selected: " + " ".join(self.files_show_names))
            # Once file is chosen "Create" and "Choose existing" buttons will be available
            self.create_table_button.configure(state="normal",
                                               relief="groove",
                                               bg='#c5eb93',
                                               border='4')

            self.update_button_text(self.files_string)

        # Adds a comma if the number of files more than one
        elif len(self.files_show_names) >= 2:
            self.files_string = ("File(s) Selected: " + ", ".join(self.files_show_names))
            max_length_allowed = 76

            # Removes the strings if the number of words exceeds the limit.
            while len(self.files_string) > max_length_allowed:
                self.files_string = self.files_string[:-1]

            # Adds  the triple dots if the number of words exceeds the limit
            if len(self.files_string) >= max_length_allowed:
                self.files_string = self.files_string + "...\n"
            # Updates the file selected text.
            self.update_button_text(self.files_string)

    def update_button_text(self, text):
        """Updates the string in the GUI"""
        self.button_text.set(text)

    def introduction_window(self):
        """Window gains information necessary information to create a payroll table from user"""
        # Reset window info
        self.files_show_names = []
        self.files_show_directory = []
        self.button_text.set("File(s) Selected: ")

        self.payroll_selection = False
        self.interface_window_remover()
        button_frame = self.introduction = Frame(self)
        button_frame.grid()

        # Short welcome text
        heading_text = ttk.Label(button_frame,
                                 text="Select one of the following:",
                                 foreground="green",
                                 font=('Arial', 21))
        # Placing coordinates
        heading_text.grid(column=0,
                          columnspan=3,
                          row=0,
                          padx=160,
                          pady=25,
                          sticky="W")
        get_previous_button = Button(button_frame,
                                     border='0',
                                     image=self.GetPreviousImage,
                                     command=self.get_table_example_window)
        get_previous_button.grid(column=0,
                                 row=4,
                                 sticky='w',
                                 padx=65)

        create_master_button = Button(button_frame,
                                      border='0',
                                      image=self.CreateMasterImage,
                                      command=self.selection_step_window)
        create_master_button.grid(column=0,
                                  row=4,
                                  sticky='w',
                                  padx=245)

        create_payroll_button = Button(button_frame,
                                       border='0',
                                       image=self.CreatePayrollImage,
                                       command=self.payroll_cost_center)
        create_payroll_button.grid(column=0,
                                   row=4,
                                   sticky='w',
                                   padx=425)

        # Button for report/request
        problem_button = Button(button_frame,
                                border='0',
                                text="Instructions & Information",
                                command=self.open_instructions_url,
                                foreground="blue",
                                font=('Arial', 11, 'underline'))
        problem_button.grid(sticky='w',
                            row=5,
                            column=0,
                            pady=30,
                            padx=25)

        # Update date text
        heading_text = Button(button_frame,
                              border='0',
                              text="Updated: 10/13/2020",
                              command=self.open_latest_release,
                              foreground="blue",
                              font=('Arial', 10, 'underline'))
        # Placing coordinates
        heading_text.grid(column=0,
                          columnspan=3,
                          row=5,
                          rowspan=6,
                          padx=495,
                          pady=0)

    def selection_step_window(self):
        # Removes any other necessary window
        if self.payroll_window is not None:
            self.payroll_selection = True
        self.interface_window_remover()

        # Creates a frame
        button_frame = self.selection_window = Frame(self)
        button_frame.grid()

        # Sets repeated text
        if self.payroll_selection is False:
            self.main_text_interface(button_frame, title_text="Master Table",
                                     back_button_function=self.introduction_window,
                                     description_text="The program will create a master table based on Excel files")

        else:
            self.main_text_interface(button_frame, title_text="Payroll Table",
                                     back_button_function=self.payroll_cost_center,
                                     description_text="The program will create a payroll table based on Excel files")

        # A button to select files
        select_files_button = Button(button_frame,
                                     relief="groove",
                                     bg='#c5eb93',
                                     border='4',
                                     text="Select all Excel files to continue",
                                     command=self.select_excel_files,
                                     foreground="green",
                                     font=('Arial', 18, 'bold'))
        select_files_button.place(x=126, y=120)

        # Sets location for files selected
        self.create_files_names.place(x=8, y=207)

        # Short description for select button
        select_files_description = tk.Label(button_frame,
                                            text='Select an excel file/files which you would '
                                                 'like to make a table from',
                                            foreground="gray",
                                            font=("Arial", 10, 'bold'))
        select_files_description.place(x=105, y=178)

        # Allows to Change/View/Delete file(s)
        modify_files_button = tk.Button(button_frame,
                                        border=0,
                                        text='Change/View/Delete file(s)',
                                        command=self.change_files_window,
                                        foreground="gray",
                                        font=("Arial", 10, "bold", 'underline'))
        modify_files_button.place(x=8, y=246)
        if self.payroll_selection is False:
            self.create_table_button = Button(button_frame,
                                              relief="groove",
                                              bg='#c5eb93',
                                              border='4',
                                              text="Create an Excel table",
                                              command=self.table_setting_window,
                                              foreground="green",
                                              font=('Arial', 16, 'bold'))
            self.create_table_button.grid(column=0,
                                          columnspan=3,
                                          row=8,
                                          pady=115,
                                          padx=400)
        else:
            self.create_table_button = Button(button_frame,
                                              relief="groove",
                                              bg='#c5eb93',
                                              border='4',
                                              text="Select folder to save",
                                              command=self.create_payroll_table,
                                              foreground="green",
                                              font=('Arial', 16, 'bold'))
            self.create_table_button.grid(column=0,
                                          columnspan=3,
                                          row=8,
                                          pady=115,
                                          padx=400)
        if not self.files_show_directory:
            # Will allow going to the next window once you selected at least one file
            self.create_table_button.configure(bg="#d9dad9",
                                               relief=SUNKEN,
                                               border='1',
                                               state="disabled")

    def delete_list_element(self):
        """Deletes selected file from a list"""

        def get_element_value(listbox):
            value = listbox.get(ACTIVE)
            return value

        def delete_list_element(listbox, list_1, list_2, v_index):
            # Removes element from two lists
            listbox.delete(ACTIVE)
            del list_1[v_index]
            list_2_value = list_2[::-1][v_index]
            list_2.remove(list_2_value)
            return list_1, list_2

        try:
            selected_file = get_element_value(self.listbox)
            if len(selected_file) == 0:
                self.selection_step_window()
            else:
                ask_message = "Would you like to delete " + selected_file + " file?"
                user_response = messagebox.askokcancel("Uni-Scheduler", ask_message)
                if user_response is True:
                    val_index = self.files_show_names.index(selected_file)
                    self.files_show_names, self.files_show_directory = delete_list_element(self.listbox,
                                                                                           self.files_show_names,
                                                                                           self.files_show_directory,
                                                                                           val_index)
                else:
                    pass
        except ValueError:
            # Returns to selection window if no files in a list
            self.selection_step_window()
        self.update_button_text(("File(s) Selected: " + " ".join(self.files_show_names)))

    def change_list_element(self):
        """Removes and adds a file"""
        get_files_amount = len(self.files_show_names)
        if get_files_amount == 0:
            # Returns to selection window if no files in a list
            self.selection_step_window()
        else:
            self.select_excel_files()

            def update_listbox(listbox, file):
                # Updates the list box to include selected file
                listbox.insert(END, file)

            if get_files_amount < len(self.files_show_names):
                self.delete_list_element()
                update_listbox(self.listbox, self.files_show_names[0])

    def change_files_window(self):
        """The window for changing/deleting selected files"""
        # Removes previous window
        self.interface_window_remover()
        button_frame = self.files_manipulation_window = Frame(self)
        button_frame.grid()
        if self.payroll_selection is False:
            self.main_text_interface(button_frame, title_text="Modify Files",
                                     back_button_function=self.selection_step_window,
                                     description_text="Change or delete the file from the current list",
                                     x_description=0)
        else:

            self.main_text_interface(button_frame, title_text="Modify Files",
                                     back_button_function=self.selection_step_window,
                                     description_text="Change or delete the file from the current list",
                                     x_description=0)

        heading_label = ttk.Label(button_frame,
                                  text="List of files selected:",
                                  foreground="green",
                                  font=('Arial', 14))
        heading_label.grid(sticky='WN',
                           column=0,
                           row=2,
                           rowspan=3,
                           padx=13,
                           pady=20)

        table_list_window = Frame(button_frame, width=300, height=100, bd=0)
        table_list_window.place(x=15, y=110)
        scrollbar = Scrollbar(table_list_window, orient=VERTICAL)

        self.listbox = Listbox(table_list_window, yscrollcommand=scrollbar.set, selectmode=SINGLE, font=0, bd=1)
        self.listbox.config(width=32, height=10)
        scrollbar.config(command=self.listbox.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.listbox.pack(side=LEFT)

        for item in self.files_show_names:
            self.listbox.insert(END, item)
        change_button = tk.Button(button_frame,
                                  text="Change file",
                                  command=self.change_list_element,
                                  foreground="green",
                                  bg='#f0f8ff',
                                  border='4',
                                  relief="groove",
                                  font=('Arial', 14))

        change_button.place(x=430, y=125)

        delete_button = tk.Button(button_frame,
                                  text="Delete file",
                                  command=self.delete_list_element,
                                  foreground="green",
                                  bg='#f0f8ff',
                                  border='4',
                                  relief="groove",
                                  font=('Arial', 14))

        delete_button.place(x=435, y=175)

        continue_button = tk.Button(button_frame,
                                    text="Continue",
                                    command=self.selection_step_window,
                                    foreground="green",
                                    bg='#c5eb93',
                                    border='4',
                                    relief="groove",
                                    font=('Arial', 14))
        continue_button.grid(column=0,
                             columnspan=2,
                             row=3,
                             rowspan=6,
                             sticky="W",
                             padx=440,
                             pady=90)

    def get_table_example_window(self):
        """Creates a table based on previous semesters web data"""
        self.interface_window_remover()

        button_frame = self.get_example_window = Frame(self)
        button_frame.grid()

        heading_label = ttk.Label(button_frame,
                                  text="Get a table by department",
                                  foreground="green",
                                  font=('Arial', 18))
        heading_label.grid(column=0,
                           row=0,
                           padx=90,
                           pady=25,
                           sticky="n")

        # Description of a reason to have this window
        description_label = ttk.Label(button_frame,
                                      text="The program will generate a table by using data from MNSU website",
                                      foreground="gray",
                                      font=('Arial', 12))

        description_label.grid(column=0,
                               row=0,
                               rowspan=2,
                               padx=85,
                               pady=75)

        # Holds variables
        variable_web_semesters = StringVar(button_frame)
        variable_web_years = StringVar(button_frame)
        variable_web_department = StringVar(button_frame)

        # Sets defaults values for interface
        variable_web_department.set(self.web_department_options[0])
        variable_web_semesters.set(self.web_semesters_options[0])
        variable_web_years.set(self.web_year_options[0])
        # Sets defaults values for script
        self.web_department_parameters = self.web_department_options[0]
        self.web_year = self.web_year_options[0]
        self.web_semester_parameters = self.web_semesters_options[0]

        department_selection_label = tk.Label(button_frame,
                                              text="Select department:",
                                              font=('Arial', 16))
        department_selection_label.place(x=30, y=160)

        selection_label = tk.Label(button_frame,
                                   text="Select the semester and year: ",
                                   font=('Arial', 16))
        selection_label.place(x=30, y=200)

        # Option menu / Check buttons
        web_department_options = OptionMenu(button_frame,
                                            variable_web_department,
                                            *self.web_department_options,
                                            command=self.return_web_department)
        web_department_options.place(x=230, y=160)
        web_department_options.configure(relief="groove",
                                         bg='#c5eb93',
                                         border='4',
                                         foreground="green",
                                         font=('Arial', 10, 'bold'))

        web_semester_options = OptionMenu(button_frame,
                                          variable_web_semesters,
                                          *self.web_semesters_options,
                                          command=self.return_web_semester)
        web_semester_options.place(x=320, y=200)
        web_semester_options.configure(relief="groove",
                                       bg='#c5eb93',
                                       border='4',
                                       foreground="green",
                                       font=('Arial', 10, 'bold'))

        web_year_options = OptionMenu(button_frame,
                                      variable_web_years,
                                      *self.web_year_options,
                                      command=self.return_web_year)
        web_year_options.place(x=425, y=200)
        web_year_options.configure(relief="groove",
                                   bg='#c5eb93',
                                   border='4',
                                   foreground="green",
                                   font=('Arial', 10, 'bold'))

        create_table_button = Button(button_frame,
                                     relief="groove",
                                     bg='#c5eb93',
                                     border='4',
                                     text="Get tables and choose folder",
                                     command=self.create_web_table,
                                     foreground="green",
                                     font=('Arial', 14))

        create_table_button.grid(sticky="E",
                                 column=0,
                                 columnspan=2,
                                 row=3,
                                 padx=20,
                                 pady=110)
        back_button = Button(button_frame,
                             border='0',
                             image=self.BackImage,
                             command=self.introduction_window)

        back_button.grid(sticky='WN',
                         column=0,
                         row=0,
                         rowspan=2,
                         pady=15,
                         padx=8)

    def table_setting_window(self):
        """Gives the ability to provide additional changes to the table if the user wants to."""

        # Removes previous window
        self.interface_window_remover()

        # Creates a new frame
        button_frame = self.settings_window = Frame(self)
        button_frame.grid()

        self.main_text_interface(button_frame, title_text="Master Table",
                                 back_button_function=self.selection_step_window,
                                 description_text="Set up the settings for a table(s).",
                                 x_description=200, y_title=28)

        # Select table type
        self.get_value = tk.IntVar()
        self.include_friday = tk.IntVar()

        # Allow the user to select the day order(incomplete)
        self.table_order_default = Radiobutton(button_frame,
                                               text="Default order",
                                               font=('Arial', 11),
                                               variable=self.get_value,
                                               command=self.user_table_choice,
                                               value=1)
        self.table_order_default.grid(column=0,
                                      row=2,
                                      sticky='sw',
                                      padx=11,
                                      pady=0)

        self.table_order_type = Radiobutton(button_frame,
                                            text="Days in order",
                                            font=('Arial', 11),
                                            variable=self.get_value,
                                            command=self.user_table_choice,
                                            value=2)
        self.table_order_type.grid(column=0,
                                   row=2,
                                   sticky='sw',
                                   padx=148,
                                   pady=0)

        out_order_button = Button(button_frame,
                                  border='0',
                                  image=self.OutOrderImage,
                                  command=self.out_order_select)
        out_order_button.grid(column=0,
                              row=3,
                              rowspan=4,
                              sticky='w',
                              padx=11,
                              pady=0)

        in_order_button = Button(button_frame,
                                 border='0',
                                 image=self.InOrderImage,
                                 command=self.in_order_select)
        in_order_button.grid(column=0,
                             row=3,
                             rowspan=4,
                             sticky='w',
                             padx=150)

        table_name_label = ttk.Label(button_frame,
                                     text="Name the table: ",
                                     foreground="green",
                                     font=('Arial', 18))
        table_name_label.grid(column=0,
                              columnspan=3,
                              row=2,
                              sticky='ES',
                              padx=80,
                              pady=10)

        # A box to allow user type a name of the table they desire
        self.table_name_insertion_box = Text(button_frame,
                                             height=1.2,
                                             width=27)
        self.table_name_insertion_box.grid(column=0,
                                           columnspan=3,
                                           row=2,
                                           rowspan=4,
                                           sticky='EN',
                                           padx=33,
                                           pady=75)

        self.table_name_insertion_box.insert(END, "    Type name here...")
        self.table_name_insertion_box.bind("<1>", self.name_of_table)
        self.table_name_insertion_box.configure(font=('Courier', 12, 'italic'),
                                                foreground="gray")
        self.table_name_insertion_box.bind("<Leave>", self.return_table_name)

        # Will provide the user with a four-year option depending on your current year.
        year_options = []
        for i in range(4):
            year_options.append(datetime.date.today().year + (i - 1))

        # Holds variables
        variable_semesters = StringVar(button_frame)
        variable_years = StringVar(button_frame)
        today_year = datetime.datetime.now()

        semesters_options = ["Fall", "Spring", "Summer 1st", "Summer 2nd", ]
        # Sets a Fall semester as a default
        variable_semesters.set(semesters_options[0])

        # Set the current year automatically
        for i in year_options:
            if today_year.year == i:
                variable_years.set(today_year.year)
            else:
                pass

        semester_label = ttk.Label(button_frame,
                                   text="Select the semester and year: ",
                                   foreground="green",
                                   font=('Arial', 16))
        semester_label.place(x=332, y=190)

        # Option menu / Check buttons
        semester_options_menu = OptionMenu(button_frame,
                                           variable_semesters,
                                           *semesters_options,
                                           command=self.return_semester)
        semester_options_menu.grid(column=0,
                                   columnspan=3,
                                   row=3,
                                   rowspan=4,
                                   sticky='e',
                                   pady=34,
                                   padx=111)
        semester_options_menu.configure(relief="groove",
                                        bg='#c5eb93',
                                        border='4',
                                        foreground="green",
                                        font=('Arial', 10, 'bold'))

        year_options_menu = OptionMenu(button_frame,
                                       variable_years,
                                       *year_options,
                                       command=self.return_year)
        year_options_menu.grid(column=0,
                               columnspan=3,
                               row=3,
                               rowspan=4,
                               sticky='e',
                               pady=34,
                               padx=37)
        year_options_menu.configure(relief="groove",
                                    bg='#c5eb93',
                                    border='4',
                                    foreground="green",
                                    font=('Arial', 10, 'bold'))

        friday_option = Checkbutton(button_frame,
                                    text="Include Friday",
                                    variable=self.include_friday,
                                    font=('Arial', 10))
        friday_option.grid(sticky="nw",
                           column=0,
                           row=7,
                           padx=7)

        next_step_button = Button(button_frame,
                                  relief="groove",
                                  bg='#c5eb93',
                                  border='4',
                                  text="Select Folder",
                                  command=self.create_master_table,
                                  foreground="green",
                                  font=('Arial', 16, 'bold'))
        next_step_button.grid(column=0,
                              columnspan=3,
                              row=5,
                              rowspan=5,
                              sticky='ENS',
                              pady=100,
                              padx=37)

    def name_of_table(self, event):
        user_input = self.table_name_insertion_box.get("1.0", END)
        if user_input[:21] == "    Type name here...":
            self.table_name_insertion_box.delete("1.0", END)
            self.table_name_insertion_box.insert(END, " ")
            self.table_name_insertion_box.configure(font=('Courier', 12, 'bold'),
                                                    foreground="black")

    def create_master_table(self):
        """Moves into the creation process"""
        self.table_friday_include = self.include_friday.get()
        self.program_loading_window(block_table=True, payroll_table=False)

    def create_web_table(self):
        self.program_loading_window(block_table=False, payroll_table=False)

    def create_payroll_table(self):
        """Moves into the creation process"""
        self.table_friday_include = 1
        self.program_loading_window(block_table=False, payroll_table=True)

    def user_table_choice(self):
        """Table days order"""
        self.table_settings_type = self.get_value.get()

    def in_order_select(self):
        """Sets variable 2 if days in order"""
        self.table_order_type.select()
        self.table_settings_type = 2

    def out_order_select(self):
        """Sets variable 1 if follows standard order"""
        self.table_order_default.select()
        self.table_settings_type = 1

    def return_year(self, year_value):
        """Captures user selection - year"""
        self.table_settings_year = year_value

    def return_web_department(self, department):
        self.web_department_parameters = department

    def return_web_year(self, year):
        self.web_year = year

    def return_web_semester(self, semester):
        self.web_semester_parameters = semester

    def return_payroll_year_1(self, year):
        self.payroll_year_1 = year

    def return_payroll_semester_1(self, semester):
        self.payroll_semester_1 = semester

    def return_payroll_year_2(self, year):
        self.payroll_year_2 = year

    def return_payroll_semester_2(self, semester):
        self.payroll_semester_2 = semester

    def return_semester(self, semester_value):
        """Captures user selection - semester"""
        self.table_settings_semester = semester_value

    def return_table_name(self, event):
        """Presents a user what he wrote as a table name"""
        self.table_settings_name = self.table_name_insertion_box.get("1.0", END)

    def return_cost_center_list(self, event):
        """Presents a user what he wrote as a table name"""
        self.cost_center_string += self.table_name_insertion_box.get("1.0", END)

    def open_master_table(self):
        """Opens a master table"""
        try:
            excel_file = self.table_settings_name.replace('\n', ' ').replace('\r', '')
            excel_file = excel_file.replace(" ", "")
            excel_file = str(excel_file) + '.' + 'xlsx'
            excel_file = os.path.join(self.folder, excel_file)
            os.startfile(excel_file)
        except FileNotFoundError:
            for filename in glob.glob(os.path.join(self.folder, '*.xlsx')):
                os.startfile(filename)

    def open_payroll_folder(self):
        os.startfile(self.folder)

    def open_excel_copies(self):
        """For error window"""
        folder_path = "copy_folder\\"
        for filename in glob.glob(os.path.join(folder_path, '*.xlsx')):
            os.startfile(filename)

    def exit_function(self):
        sys.exit()

    def program_loading_window(self, block_table=True, payroll_table=False, count=0):
        self.interface_window_remover()

        button_frame = self.creating_step_window = Frame(self)
        button_frame.grid()

        self.folder = filedialog.askdirectory(title='Please select a directory')
        if self.folder == "":
            if count == 1:
                self.introduction_window()
            else:
                self.program_loading_window(block_table, payroll_table, count=1)

        else:
            def create_web_table(web_department_parameters, urlencode_dict_list, web_semester_parameters, web_year,
                                 web_department_options, folder):
                """Department chairs might need an example of a file from the previous semester.
                This function will create a table based on university records."""
                urlencode_list = []

                folder_path = folder

                def create_table(urlencode_dict, web_department, web_semester, c_web_year, get_all_tables=False):
                    if not urlencode_dict:
                        previous_data.PreviousCourses(folder_path, web_department, web_semester, int(c_web_year),
                                                      get_all=get_all_tables)
                    else:
                        for len_list in range(len(urlencode_dict)):
                            for departament_semester, urlencode in urlencode_dict[len_list].items():
                                if departament_semester == web_department:
                                    urlencode_list.append(urlencode)
                                if departament_semester == web_semester:
                                    urlencode_list.append(urlencode)
                        if len(urlencode_list) == 2:
                            if get_all_tables is True:
                                previous_data.PreviousCourses(folder_path, web_department, web_semester,
                                                              int(c_web_year), urlencode_list[0], urlencode_list[1],
                                                              get_all=get_all_tables)
                            else:
                                previous_data.PreviousCourses(folder_path, web_department, web_semester,
                                                              int(c_web_year), urlencode_list[0], urlencode_list[1],
                                                              get_all=get_all_tables)
                        else:
                            if get_all_tables is True:
                                previous_data.PreviousCourses(folder_path, web_department, web_semester,
                                                              int(c_web_year), get_all=get_all_tables)
                            else:
                                previous_data.PreviousCourses(folder_path, web_department, web_semester,
                                                              int(c_web_year), get_all=get_all_tables)

                if web_department_parameters not in ["All COB Departments", "ACCT & BLAW & MACC",
                                                     "MRKT & IBUS", "MGMT & MBA"]:
                    try:
                        create_table(urlencode_dict_list, web_department_parameters,
                                     web_semester_parameters, web_year)
                        if os.path.isdir(folder_path):
                            os.startfile(folder_path)
                    except PermissionError:
                        messagebox.showwarning("Existing excel file open!",
                                               "Please close your current excel files and try again.")
                else:
                    if web_department_parameters == "All COB Departments":
                        web_department_options.remove('MGMT & MBA')
                        web_department_options.remove('MRKT & IBUS')
                        web_department_options.remove('ACCT & BLAW & MACC')
                        web_department_options.remove('All COB Departments')
                        dep = iter(web_department_options)
                        for department in dep:
                            create_table(urlencode_dict_list, department,
                                         web_semester_parameters, web_year, get_all_tables=True)
                    elif web_department_parameters == "ACCT & BLAW & MACC":
                        create_table(urlencode_dict_list, "ACCT",
                                     web_semester_parameters, web_year)
                        create_table(urlencode_dict_list, "BLAW",
                                     web_semester_parameters, web_year)
                        create_table(urlencode_dict_list, "MACC",
                                     web_semester_parameters, web_year)
                    elif web_department_parameters == "MRKT & IBUS":
                        create_table(urlencode_dict_list, "MRKT",
                                     web_semester_parameters, web_year)
                        create_table(urlencode_dict_list, "IBUS",
                                     web_semester_parameters, web_year)
                    elif web_department_parameters == "MGMT & MBA":
                        create_table(urlencode_dict_list, "MGMT",
                                     web_semester_parameters, web_year)
                        create_table(urlencode_dict_list, "MBA",
                                     web_semester_parameters, web_year)
                    else:
                        pass
                    if os.path.isdir(folder_path):
                        os.startfile(folder_path)

            global switch
            switch = False

            def processor():
                global switch

                wait_text = StringVar()
                while not switch:
                    wait_label = tk.Label(button_frame, textvariable=wait_text,
                                          foreground="green",
                                          font=('Courier', 20, 'bold'))
                    wait_label.grid(column=1, row=2, rowspan=2, padx=10, pady=10)
                    if block_table is True:
                        wait_text.set("\r \n \n \n  Creating a scheduling table...")
                    elif payroll_table is True:
                        wait_text.set("\r \n \n \n  Creating a Payroll table...")
                    else:
                        wait_text.set("\r \n \n \n  Creating a table from web...")

                    wait_label.configure(textvariable=wait_text)

                    sys.stdout.flush()
                    time.sleep(0.1)

            processor_threading = threading.Thread(target=processor, name="processor thread")
            processor_threading.start()
            button_frame.update()

            # Moves to the next class which is processing all the files
            # try:
            if block_table is True:
                self.error_data_list = receiver.DataProcessor(self.folder,
                                                              self.files_show_directory, self.table_settings_name,
                                                              self.table_settings_semester, self.table_settings_year,
                                                              self.table_settings_type,
                                                              self.table_friday_include,
                                                              self.room_cap_dict, payroll_table).get_excel_errors()

                self.user_result_window()

            elif payroll_table is False:
                create_web_table(self.web_department_parameters, self.urlencode_dict_list,
                                 self.web_semester_parameters, self.web_year, self.web_department_options,
                                 folder=self.folder)
                self.introduction_window()
            else:
                self.error_data_list = receiver.DataProcessor(self.folder, self.files_show_directory,
                                                              self.table_settings_name,
                                                              self.table_settings_semester,
                                                              self.table_settings_year,
                                                              self.table_settings_type,
                                                              self.table_friday_include,
                                                              self.room_cap_dict, payroll_table)

                self.payroll_finish_window()
            """
            except Exception as e:
                tk.messagebox.showerror(title="Program failed", message="Program failed... Please try again.")
                self.introduction_window()

            switch = True
            """

    def user_result_window(self):
        self.interface_window_remover()

        def clear_error_data(error_data):
            """Clears all unnecessary errors."""
            clear_data_list = []
            clear_data_dict = {}
            for i in range(len(error_data)):
                if error_data[i - 1].get("Color") == 'FF687B' or error_data[i - 1].get("Color") == 'FEBBBB':
                    try:
                        if error_data[i - 1].get("Comment") == error_data[i].get("Comment"):
                            pass
                        else:
                            clear_data_dict['Message'] = error_data[i - 1].get("Comment")
                            clear_data_list.append(clear_data_dict.copy())
                    except IndexError:
                        pass
            return clear_data_list

        def remove_dict_duplicates(error_data):
            new_dict_list = []
            for i in range(len(error_data)):
                if error_data[i] not in error_data[i + 1:]:
                    new_dict_list.append(error_data[i])
            return new_dict_list

        if self.error_data_list == 'User_Doesnt_Listen':
            self.interface_window_remover()
            self.selection_step_window()

        clear_error_list = clear_error_data(self.error_data_list)
        clear_error_list = remove_dict_duplicates(clear_error_list)
        if len(clear_error_list) != 0:

            button_frame = self.notification_window = Frame(self)
            button_frame.grid()

            self.main_text_interface(button_frame, title_text="Master Table",
                                     back_button_function=self.selection_step_window,
                                     remove_back=True)

            total_number_mistakes = str(int(len(clear_error_list) / 2))
            if total_number_mistakes == '0':
                total_number_mistakes = '1'
            total_number_mistakes = 'Possible mistakes: ' + total_number_mistakes
            red_error_label = tk.Label(button_frame,
                                       text=total_number_mistakes,
                                       foreground="gray",
                                       font=('Courier', 14, 'bold'))
            red_error_label.grid(sticky='w',
                                 column=0,
                                 columnspan=4,
                                 row=2,
                                 padx=10)

            # Creating a scroll window of errors
            def scroll_error_messages():
                """Shows all the errors"""
                for error_len in range(len(clear_error_list)):
                    if str(clear_error_list[error_len].get("Message")) == "None":
                        pass
                    elif str(clear_error_list[error_len].get(
                            "Message")) == "A program couldn't read this row correctly. " \
                                           "Report it if needed.":
                        ui_message = 'Check for additional errors by pressing "Open excel copies"' + (' ' * 100)
                        Label(frame, text=ui_message, background="#ee8282").grid(sticky="w", row=99, column=0)
                    else:
                        # Shows only even to eliminate repetitive conflicts
                        if error_len % 2 == 0:
                            conflict_row_message = str(clear_error_list[error_len].get("Message"))
                            Label(frame, text=conflict_row_message, background="#ee8282").grid(sticky="w",
                                                                                               row=error_len, column=0)

            def show_all_messages(event):
                canvas.configure(scrollregion=canvas.bbox("all"), width=600, height=60)

            error_message_frame = Frame(button_frame, relief=GROOVE, width=600, height=150, bd=1)
            error_message_frame.place(x=22, y=100)

            canvas = Canvas(error_message_frame)
            frame = Frame(canvas)

            # Scroll bar on a right side
            user_scrollbar_y = tk.Scrollbar(error_message_frame, orient="vertical")
            user_scrollbar_y.pack(side=RIGHT, fill=Y)
            canvas.configure(yscrollcommand=user_scrollbar_y.set)
            canvas.pack(side=RIGHT, fill=BOTH)
            user_scrollbar_y.config(command=canvas.yview)

            canvas.create_window((0, 0), window=frame, anchor='nw')
            frame.bind("<Configure>", show_all_messages)
            frame.bind("<Enter>", show_all_messages)
            frame.bind("<Leave>", show_all_messages)

            instructions_message = 'Use "Open excel copies" button to check if the program ' \
                                   'found any conflicts or missing information.'

            instructions_message_label = ttk.Label(button_frame,
                                                   text=instructions_message,
                                                   foreground="gray",
                                                   font=('Arial', 10, 'bold'))
            instructions_message_label.grid(sticky="W",
                                            column=0,
                                            columnspan=3,
                                            row=3,
                                            padx=13,
                                            pady=31)
            scroll_error_messages()

            open_copies_button = Button(button_frame,
                                        border='0',
                                        image=self.ExcelCopyFile,
                                        command=self.open_excel_copies)
            open_copies_button.grid(sticky="NW",
                                    column=0,
                                    row=3,
                                    rowspan=4,
                                    padx=20,
                                    pady=90)

            open_copies_text = Button(button_frame,
                                      border='0',
                                      text="Open excel copies",
                                      command=self.open_excel_copies,
                                      foreground="gray",
                                      font=('Arial', 11, 'bold'))
            open_copies_text.grid(sticky="W",
                                  column=0,
                                  row=3,
                                  rowspan=4,
                                  padx=8,
                                  pady=193)

            open_main_button = Button(button_frame,
                                      border='0',
                                      image=self.ExcelMainFile,
                                      command=self.open_master_table)
            open_main_button.grid(sticky="N",
                                  column=0,
                                  columnspan=2,
                                  row=3,
                                  rowspan=4,
                                  pady=90)

            open_main_text = Button(button_frame,
                                    border='0',
                                    text="Open master table",
                                    command=self.open_master_table,
                                    foreground="gray",
                                    font=('Arial', 11, 'bold'))
            open_main_text.grid(sticky="WE",
                                column=0,
                                columnspan=3,
                                row=3,
                                rowspan=4,
                                padx=150,
                                pady=0)

            exit_button = Button(button_frame,
                                 border='0',
                                 image=self.ExitApplicationImage,
                                 command=self.exit_function)
            exit_button.grid(sticky="NE",
                             column=0,
                             columnspan=3,
                             row=3,
                             rowspan=4,
                             padx=25,
                             pady=90)

            exit_text = Button(button_frame,
                               border='0',
                               text="Exit",
                               command=self.exit_function,
                               foreground="gray",
                               font=('Arial', 11, 'bold'))
            exit_text.grid(sticky="E",
                           column=0,
                           columnspan=4,
                           row=3,
                           rowspan=4,
                           padx=63,
                           pady=0)
        else:

            button_frame = self.notification_window = Frame(self)
            button_frame.grid()

            instructions_message = "Everything looks great! ✔"

            no_errors_label = ttk.Label(button_frame,
                                        text=instructions_message,
                                        foreground="green",
                                        font=('Arial', 24, 'bold'))
            no_errors_label.grid(column=0,
                                 columnspan=3,
                                 row=1,
                                 padx=130,
                                 pady=30)

            open_file_button = Button(button_frame,
                                      relief="groove",
                                      bg='#c5eb93',
                                      border='4',
                                      text="Open Master Table",
                                      command=self.open_master_table,
                                      foreground="green",
                                      font=('Arial', 20, 'bold'))
            open_file_button.grid(columnspan=3,
                                  row=2,
                                  pady=40)

            exit_program_button = Button(button_frame,
                                         relief="groove",
                                         bg='#c5eb93',
                                         border='4',
                                         text="Exit",
                                         command=self.exit_function,
                                         foreground="green",
                                         font=('Arial', 16, 'bold'))
            exit_program_button.grid(columnspan=3,
                                     row=2,
                                     rowspan=3,
                                     pady=65)

            user_feedback_button = Button(button_frame,
                                          border='0',
                                          text="Please provide feedback about your experience.",
                                          command=self.submit_ticket_form,
                                          foreground="blue",
                                          font=('Arial', 11, 'underline'))
            user_feedback_button.grid(column=0,
                                      columnspan=3,
                                      row=3,
                                      sticky="EW",
                                      pady=60)

    def payroll_cost_center(self):
        self.interface_window_remover()

        button_frame = self.payroll_window = Frame(self)
        button_frame.grid()

        self.main_text_interface(button_frame, title_text="Payroll Table",
                                 description_text="Provide a cost center for each department for a payroll table",
                                 back_button_function=self.introduction_window, x_description=110)

        comma_note = ttk.Label(button_frame,
                               text="- Use a comma if the specific department has multiple cost\n centers.",
                               foreground="green",
                               font=('Arial', 11, 'bold'))
        comma_note.place(x=200, y=125)

        prof_note = ttk.Label(button_frame,
                              text='- Type "Professor" if the department cost center is based on\n '
                                   'the professor of other departments',
                              foreground="green",
                              font=('Arial', 11, 'bold'))
        prof_note.place(x=200, y=175)

        example_note = ttk.Label(button_frame,
                                 text="Example: 'BUS => Professor' will result in giving each\n"
                                      "faculty department of cost center",
                                 foreground="gray",
                                 font=('Arial', 11, 'bold'))
        example_note.place(x=200, y=233)

        self.move_next_step = Button(button_frame,
                                     relief="groove",
                                     bg='#c5eb93',
                                     border='4',
                                     text="Next step >",
                                     command=self.selection_step_window,
                                     foreground="green",
                                     font=('Arial', 16, 'bold'))
        self.move_next_step.grid(sticky='e',
                                 column=0,
                                 columnspan=2,
                                 row=6,
                                 pady=110,
                                 padx=20)

        def show_all_departments(event):
            canvas.configure(scrollregion=canvas.bbox("all"), width=125, height=190)

        self.cost_department_list = Frame(button_frame, relief=GROOVE, width=125, height=190, bd=1)
        self.cost_department_list.grid()
        self.cost_department_list.place(x=40, y=110)

        canvas = Canvas(self.cost_department_list)

        self.mini_frame = Frame(canvas)

        # Scroll bar on a right side
        user_scrollbar_y = tk.Scrollbar(self.cost_department_list, orient="vertical")
        user_scrollbar_y.pack(side=RIGHT, fill=Y)
        canvas.configure(yscrollcommand=user_scrollbar_y.set)
        canvas.pack(side=RIGHT, fill=BOTH)
        user_scrollbar_y.config(command=canvas.yview)

        canvas.create_window((0, 0), window=self.mini_frame, anchor='nw')
        self.mini_frame.bind("<Configure>", show_all_departments)
        self.mini_frame.bind("<Enter>", show_all_departments)
        self.mini_frame.bind("<Leave>", show_all_departments)

        self.scroll_error_messages()

    def scroll_error_messages(self):
        """Shows all the errors"""

        def get_csv_file(file):
            cost_center = dict()
            if os.path.isfile(file):
                with open(file) as csv_file:
                    read_csv_file = csv.DictReader(csv_file, delimiter=',')
                    for row in read_csv_file:
                        cost_center = dict(row)
                return cost_center

        csv_file_data = get_csv_file(f'{self.cwd}\\department_cost.csv')

        cob_department_list = ["Marketing & International Business", "Accounting", "Business Law", "Finance",
                               "MACC", "Management", "MBA", "BUS"]

        self.cost_box_insert = []
        department_label_list = []

        for i in range(len(cob_department_list)):
            if cob_department_list[i] == "Marketing & International Business":
                short_abbrev = "Marketing & I. Business"
                department_label_list.append(Label(self.mini_frame, text=short_abbrev))  # creates entry boxes
            else:
                department_label_list.append(Label(self.mini_frame, text=cob_department_list[i]))
            self.cost_box_insert.append(Entry(self.mini_frame, text=cob_department_list[i]))  # creates entry boxes
            department_label_list[i].pack()
            if csv_file_data is None:
                pass
            else:
                self.cost_box_insert[i].delete(0, 'end')  # Clearing entry box
                self.cost_box_insert[i].insert(END, csv_file_data.get(cob_department_list[i]))

            self.cost_box_insert[i].pack()

    def cost_dict(self):
        for i in range(len(self.cost_box_insert)):
            self.department_cost_dict.update({self.cost_box_insert[i].cget("text"): self.cost_box_insert[i].get()})

        # Writes a csv file
        cost_file = f'{self.cwd}\\department_cost.csv'
        try:
            with open(cost_file, 'w') as new_file:
                write_file = csv.DictWriter(new_file, self.department_cost_dict.keys())
                write_file.writeheader()
                write_file.writerow(self.department_cost_dict)
        except PermissionError:
            tk.messagebox.showerror("Please close .csv file")
            self.introduction_window()

    def payroll_finish_window(self):
        self.interface_window_remover()

        button_frame = self.payroll_window = Frame(self)
        button_frame.grid()

        instructions_message = "Payroll table(s) has been created ✔"

        table_created_text = ttk.Label(button_frame,
                                       text=instructions_message,
                                       foreground="green",
                                       font=('Arial', 24, 'bold'))
        table_created_text.grid(column=0,
                                columnspan=3,
                                row=1,
                                rowspan=2,
                                padx=55,
                                pady=30)

        possible_error_notification = ttk.Label(button_frame,
                                                text="Please check for error at the end of excel file",
                                                foreground="green",
                                                font=('Arial', 12))
        possible_error_notification.grid(column=0,
                                         sticky='s',
                                         columnspan=3,
                                         row=2,
                                         padx=150,
                                         pady=0)

        open_file_button = Button(button_frame,
                                  relief="groove",
                                  bg='#c5eb93',
                                  border='4',
                                  text="Open a Folder",
                                  command=self.open_payroll_folder,
                                  foreground="green",
                                  font=('Arial', 20, 'bold'))
        open_file_button.grid(columnspan=3,
                              row=3,
                              pady=30)

        exit_program_button = Button(button_frame,
                                     relief="groove",
                                     bg='#c5eb93',
                                     border='4',
                                     text="Exit",
                                     command=self.exit_function,
                                     foreground="green",
                                     font=('Arial', 16, 'bold'))
        exit_program_button.grid(columnspan=3,
                                 row=3,
                                 rowspan=3,
                                 pady=150)

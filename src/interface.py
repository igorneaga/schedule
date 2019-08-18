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

from src import receiver, previous_semesters, previous_data


class UserInterface(Frame):
    # It is better to define values like the following ones as constants (uppercase) in a single place (like here)
    GOOGLE_FORM_URL = 'https://goo.gl/forms/wNkzjymOQ7wiNavf1'
    INSTRUCTIONS_URL = 'https://docs.google.com/document/d/1htRsKmxDX33yawrYqeHkCLWlEL-juRjeM-if8N4f2yo/edit?usp=sharing'

    def __init__(self, master):
        super().__init__(master)
        self.grid()

        # Assets
        cwd = os.getcwd()
        self.FixImage = tk.PhotoImage(file=cwd + '\\src\\assets\\report_x45.png')
        self.InfoImage = tk.PhotoImage(file=cwd + '\\src\\assets\\info_x45.png')
        self.BackImage = tk.PhotoImage(file=cwd + '\\src\\assets\\back_icon_45x45.png')
        self.OutOrderImage = tk.PhotoImage(file=cwd + '\\src\\assets\\table_v05_default.png')
        self.InOrderImage = tk.PhotoImage(file=cwd + '\\src\\assets\\table_v05_in_order.png')
        self.ExcelCopyFile = tk.PhotoImage(file=cwd + '\\src\\assets\\excel_files_icon.png')
        self.ExcelMainFile = tk.PhotoImage(file=cwd + '\\src\\assets\\master_file_icon.png')
        self.CreateMasterImage = tk.PhotoImage(file=cwd + '\\src\\assets\\create_master.png')
        self.CreatePayrollImage = tk.PhotoImage(file=cwd + '\\src\\assets\\create_fwm_table2.png')
        self.GetPreviousImage = tk.PhotoImage(file=cwd + '\\src\\assets\\get_prev_tables.png')
        self.ExitApplicationImage = tk.PhotoImage(file=cwd + '\\src\\assets\\quit_button.png')
        self.ApplicationLogoImage = tk.PhotoImage(file=cwd + '\\src\\assets\\u_logo.png')

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

        # Stores data about room capacity
        self.room_cap_dict = dict()

        # A label which will keep updating once user choose a data file
        self.button_text = tk.StringVar()
        self.button_text.set("File(s) Selected: ")
        self.create_files_names = Button(self.selection_window, border=0,
                                         textvariable=self.button_text, command=self.change_files_window,
                                         foreground="gray", font=("Arial", 11, "bold"))

        # User directory shortcut
        self.user_directory = "/"

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

        try:
            excel_file = glob.glob('__excel_files/*.xlsx')
            if not excel_file:
                pass
            else:
                open(excel_file[0], "r+")
        except IOError:
            messagebox.showerror("Close File", "Please close excel files to eliminate errors")

        self.introduction_window()

    def submit_ticket_form(self):
        """Opens a Google Form to collect any reports or requests"""
        webbrowser.open(self.GOOGLE_FORM_URL)

    def open_instructions_url(self):
        """Instructions on how to use this program"""
        webbrowser.open(self.INSTRUCTIONS_URL)

    def main_text_interface(self, button_frame, x=52, include_instructions=True):
        """Repeated title text"""
        title_label = ttk.Label(button_frame,
                                text="Schedule Builder",
                                foreground="green",
                                font=('Courier', 20, 'bold'))
        title_label.grid(sticky='W',
                         column=0,
                         row=0,
                         rowspan=2,
                         padx=x,
                         pady=10)

        # Short descriptions
        app_description_label = ttk.Label(button_frame, text="Creates a room table for courses...    ",
                                          foreground="gray",
                                          font=("Courier", 10, 'bold'))
        app_description_label.grid(sticky="SW",
                                   column=0,
                                   row=1,
                                   padx=x)

        # Button for report/request
        ticket_form_button = Button(button_frame,
                                    border='0',
                                    text="Report a problem or request",
                                    command=self.submit_ticket_form,
                                    foreground="blue",
                                    font=('Arial', 11, 'underline'))
        ticket_form_button.grid(sticky="NE",
                                column=2,
                                row=1,
                                pady=2,
                                padx=3)

        if include_instructions is True:
            # Button for instructions
            info_button = Button(button_frame,
                                 border='0',
                                 text="Instructions/Information",
                                 foreground="blue",
                                 command=self.open_instructions_url,
                                 font=('Arial', 11, 'underline'))
            info_button.grid(sticky="SE",
                             column=2,
                             row=2,
                             padx=3,
                             pady=0)
            info_button.config(image=self.InfoImage,
                               compound=RIGHT)

        ticket_form_button.config(image=self.FixImage,
                                  compound=RIGHT)

    def interface_window_remover(self):
        """Removes window once a user goes to a next step or previous step."""

        if self.introduction:
            self.introduction.grid_remove()

        if self.payroll_window:
            self.payroll_window.grid_remove()
            self.cost_department_list.grid_remove()

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
        symbol = '\ '
        self.files_show_names = []
        for i in self.files_show_directory:
            z = 0
            for j in i:
                z -= 1
                if i[z] == '/' or i[z] == symbol[0]:
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

            # Removes the strings if the number of words exceeds the limit.
            while len(self.files_string) > 83:
                self.files_string = self.files_string[:-1]

            # Adds  the triple dots if the number of words exceeds the limit
            if len(self.files_string) >= 83:
                self.files_string = self.files_string + "...\n"
            # Updates the file selected text.
            self.update_button_text(self.files_string)

    def update_button_text(self, text):
        """Updates the string in the GUI"""
        self.button_text.set(text)

    def introduction_window(self):
        """Window gains information necessary information to create a payroll table from user"""

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
                                       command=self.payroll_table_first_step)
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
                            padx=245)

    def payroll_table_first_step(self):
        pass

    def selection_step_window(self):
        # Removes any other necessary window
        self.interface_window_remover()

        # Creates a frame
        button_frame = self.selection_window = Frame(self)
        button_frame.grid()

        # Sets repeated text
        self.main_text_interface(button_frame)

        # The back button which will allow moving to the previous window
        back_button = Button(button_frame,
                             border='0',
                             image=self.BackImage,
                             command=self.introduction_window)
        back_button.grid(sticky='WN',
                         column=0,
                         row=1,
                         rowspan=2,
                         pady=7,
                         padx=3)

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

        self.create_table_button = Button(button_frame,
                                          relief="groove",
                                          bg='#c5eb93',
                                          border='4',
                                          text="Create an Excel table",
                                          command=self.table_setting_window,
                                          foreground="green",
                                          font=('Arial', 16, 'bold'))
        self.create_table_button.grid(column=1,
                                      columnspan=2,
                                      row=8,
                                      pady=188,
                                      padx=0)

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

        self.main_text_interface(button_frame)

        # The back button which will allow moving to the previous window
        back_button = Button(button_frame,
                             border='0',
                             image=self.BackImage,
                             command=self.selection_step_window)
        back_button.grid(sticky='WN',
                         column=0,
                         row=1,
                         rowspan=2,
                         pady=7,
                         padx=3)

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
        continue_button.grid(column=1,
                             columnspan=2,
                             row=3,
                             rowspan=6,
                             sticky="WS",
                             padx=20,
                             pady=135)

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
                           padx=25,
                           pady=25,
                           sticky="n")

        # Logo of a program on the right
        application_logo = Label(button_frame, image=self.ApplicationLogoImage)
        application_logo.photo = self.ApplicationLogoImage
        application_logo.grid(column=1,
                              row=0,
                              sticky="w",
                              padx=0,
                              pady=0)
        # Description of a reason to have this window
        description_label = ttk.Label(button_frame,
                                      text="The program will generate a table by using data from MNSU website",
                                      foreground="gray",
                                      font=('Arial', 12))

        description_label.grid(column=0,
                               row=0,
                               rowspan=2,
                               padx=20,
                               pady=75)
        # Semesters options
        web_semesters_options = []
        web_department_options = []
        web_year_options = []

        for param_len in range(len(self.param)):
            # Finds available options from scraping
            for key in self.param[param_len]:
                test_dict = dict()
                if key[0:4] == "FALL" or key[0:4] == "SPRI":
                    find_year_index = key.find("2")
                    web_semesters_options.append(key[0:find_year_index])
                    web_year_options.append(key[find_year_index:])
                    test_dict[key[0:find_year_index]] = self.param[param_len].get(key)
                    self.urlencode_dict_list.append(test_dict)

                else:
                    symbol_index = key.find("(")
                    web_department_options.append(key[symbol_index + 1:-1])
                    test_dict[key[symbol_index + 1:-1]] = self.param[param_len].get(key)
                    self.urlencode_dict_list.append(test_dict)

        # Holds variables
        variable_web_semesters = StringVar(button_frame)
        variable_web_years = StringVar(button_frame)
        variable_web_department = StringVar(button_frame)
        # Sets defaults values for interface
        variable_web_department.set(web_department_options[0])
        variable_web_semesters.set(web_semesters_options[0])
        variable_web_years.set(web_year_options[0])
        # Sets defaults values for script
        self.web_department_parameters = web_department_options[0]
        self.web_year = web_year_options[0]
        self.web_semester_parameters = web_semesters_options[0]

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
                                            *web_department_options,
                                            command=self.return_web_department)
        web_department_options.place(x=230, y=160)
        web_department_options.configure(relief="groove",
                                         bg='#c5eb93',
                                         border='4',
                                         foreground="green",
                                         font=('Arial', 10, 'bold'))

        web_semester_options = OptionMenu(button_frame,
                                          variable_web_semesters,
                                          *web_semesters_options,
                                          command=self.return_web_semester)
        web_semester_options.place(x=320, y=200)
        web_semester_options.configure(relief="groove",
                                       bg='#c5eb93',
                                       border='4',
                                       foreground="green",
                                       font=('Arial', 10, 'bold'))

        web_year_options = OptionMenu(button_frame,
                                      variable_web_years,
                                      *web_year_options,
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
                                     text="Create table",
                                     command=self.create_web_table,
                                     foreground="green",
                                     font=('Arial', 16, 'bold'))
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

    def create_web_table(self):
        """Department chairs might need an example of a file from the previous semester. This function will create a
        table based on university records."""
        urlencode_list = []

        try:
            if not self.urlencode_dict_list:
                previous_data.PreviousCourses(self.web_department_parameters, self.web_semester_parameters,
                                              int(self.web_year))
            else:
                for len_list in range(len(self.urlencode_dict_list)):
                    for departament_semester, urlencode in self.urlencode_dict_list[len_list].items():
                        if departament_semester == self.web_department_parameters:
                            urlencode_list.append(urlencode)
                        if departament_semester == self.web_semester_parameters:
                            urlencode_list.append(urlencode)
                if len(urlencode_list) == 2:
                    previous_data.PreviousCourses(self.web_department_parameters, self.web_semester_parameters,
                                                  int(self.web_year),
                                                  urlencode_list[0], urlencode_list[1])
                else:
                    previous_data.PreviousCourses(self.web_department_parameters, self.web_semester_parameters,
                                                  int(self.web_year))

            if os.path.isdir('web_files\\'):
                os.startfile('web_files\\' + self.web_department_parameters + "_" + str(self.web_year) + ".xlsx")
                # Gives some time to launch the excel file
                time.sleep(1)
            else:
                messagebox.showinfo("Error occurred", "Something went wrong... Try again")
            # Brings back to the main menu
            self.introduction_window()
        except PermissionError:
            messagebox.showwarning("Existing excel file open!", "Please close your current excel files and try again.")
            self.get_table_example_window()

    def table_setting_window(self):
        """Gives the ability to provide additional changes to the table if the user wants to."""

        # Removes previous window
        self.interface_window_remover()

        # Creates a new frame
        button_frame = self.settings_window = Frame(self)
        button_frame.grid()

        self.main_text_interface(button_frame)

        # The back button which will allow moving to the previous window
        back_button = Button(button_frame,
                             border='0',
                             image=self.BackImage,
                             command=self.selection_step_window)
        back_button.grid(sticky='WN',
                         column=0,
                         row=1,
                         rowspan=2,
                         pady=7,
                         padx=3)

        # Empty row in GUI for design purpose
        empty_row = ttk.Label(button_frame,
                              text=" ")
        empty_row.grid(columnspan=3,
                       row=3)

        heading_label = ttk.Label(button_frame,
                                  text="Table settings:",
                                  foreground="green",
                                  font=('Arial', 14))
        heading_label.grid(sticky='WN',
                           column=0,
                           row=2,
                           rowspan=3,
                           padx=13,
                           pady=20)

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
                                      row=3,
                                      sticky='w',
                                      padx=11)

        self.table_order_type = Radiobutton(button_frame,
                                            text="Days in order",
                                            font=('Arial', 11),
                                            variable=self.get_value,
                                            command=self.user_table_choice,
                                            value=2)
        self.table_order_type.grid(column=0,
                                   row=3,
                                   sticky='w',
                                   padx=148)

        out_order_button = Button(button_frame,
                                  border='0',
                                  image=self.OutOrderImage,
                                  command=self.out_order_select)
        out_order_button.grid(column=0,
                              row=4,
                              sticky='w',
                              padx=11)

        in_order_button = Button(button_frame,
                                 border='0',
                                 image=self.InOrderImage,
                                 command=self.in_order_select)
        in_order_button.grid(column=0,
                             row=4,
                             sticky='w',
                             padx=150)

        table_name_label = ttk.Label(button_frame,
                                     text="Name the table: ",
                                     foreground="green",
                                     font=('Arial', 18))
        table_name_label.grid(column=0,
                              columnspan=3,
                              row=3,
                              sticky='ES',
                              padx=95,
                              pady=10)

        # A box to allow user type a name of the table they desire
        self.table_name_insertion_box = Text(button_frame,
                                             height=1.2,
                                             width=27)
        self.table_name_insertion_box.grid(column=0,
                                           columnspan=3,
                                           row=3,
                                           rowspan=4,
                                           sticky='EN',
                                           padx=49,
                                           pady=50)

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

        semesters_options = ["Fall", "Spring"]
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
                                   row=4,
                                   rowspan=5,
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
                               row=4,
                               rowspan=5,
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
                           row=8,
                           padx=7)

        next_step_button = Button(button_frame,
                                  relief="groove",
                                  bg='#c5eb93',
                                  border='4',
                                  text="Create table",
                                  command=self.create_master_table,
                                  foreground="green",
                                  font=('Arial', 16, 'bold'))
        next_step_button.grid(column=0,
                              columnspan=3,
                              row=5,
                              rowspan=5,
                              sticky='EN',
                              pady=0,
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
        self.program_loading_window()

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
            self.table_settings_name.replace(" ", "")
            os.startfile('__excel_files\\ \n.xlsx')
        except FileNotFoundError:
            for filename in glob.glob(os.path.join('__excel_files\\', '*.xlsx')):
                os.startfile(filename)

    def open_excel_copies(self):
        folder_path = "copy_folder\\"
        for filename in glob.glob(os.path.join(folder_path, '*.xlsx')):
            os.startfile(filename)

    def exit_function(self):
        sys.exit()

    def program_loading_window(self, block_table=True, gains_data=False, payroll_table=False):
        self.interface_window_remover()

        button_frame = self.creating_step_window = Frame(self)
        button_frame.grid()

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
                wait_label.configure(textvariable=wait_text)

                sys.stdout.flush()
                time.sleep(0.1)

        processor_threading = threading.Thread(target=processor, name="processor thread")
        processor_threading.start()
        button_frame.update()

        # Moves to the next class which is processing all the files
        if block_table is True:
            self.error_data_list = receiver.DataProcessor(self.files_show_directory, self.table_settings_name,
                                                          self.table_settings_semester, self.table_settings_year,
                                                          self.table_settings_type,
                                                          self.table_friday_include,
                                                          self.room_cap_dict).get_excel_errors()
        switch = True
        self.user_result_window()

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

        if self.error_data_list == 'User_Doesnt_Listen':
            self.interface_window_remover()
            self.selection_step_window()
        clear_error_list = clear_error_data(self.error_data_list)
        if len(clear_error_list) != 0:

            button_frame = self.notification_window = Frame(self)
            button_frame.grid()

            self.main_text_interface(button_frame, x=10)

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
            error_message_frame.place(x=22, y=110)

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
            open_main_button.grid(sticky="NE",
                                  column=0,
                                  columnspan=2,
                                  row=3,
                                  rowspan=4,
                                  padx=10,
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
                                padx=160,
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

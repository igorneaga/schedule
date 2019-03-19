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
from tkinter import ttk

import receiver


# Temporary


class UserInterface(Frame):

    def __init__(self, master):
        super().__init__(master)
        self.grid()

        # Assets
        self.fix_image = tk.PhotoImage(file='assets\\back_icon_45x45.png')
        self.info_image = tk.PhotoImage(file='assets\\back_icon_45x45.png')
        self.back_image = tk.PhotoImage(file='assets\\back_icon_45x45.png')
        self.out_order_image = tk.PhotoImage(file='assets\\table_default_order.png')
        self.in_order_image = tk.PhotoImage(file='assets\\table_in_order.png')
        self.excel_copy_fie = tk.PhotoImage(file='assets\\excel_files_icon.png')
        self.excel_main_file = tk.PhotoImage(file='assets\\master_file_icon.png')
        self.exit_file = tk.PhotoImage(file='assets\\quit_button.png')

        self.out_order_image = self.out_order_image.subsample(2)
        self.in_order_image = self.in_order_image.subsample(2)

        # Default table characteristics
        self.table_settings_type = 1
        self.table_settings_year = "2019"
        self.table_settings_semester = "Fall"
        self.table_settings_name = "Uni_Table"
        self.table_friday_include = 0

        # Information from data files
        self.file_name = None
        self.files_show_directory = []
        self.files_show_names = []
        self.files_string = None

        # GUI windows
        self.first_step_window = None
        self.second_step_window = None
        self.creating_step_window = None
        self.settings_window = None
        self.notification_window = None

        # GUI buttons, radio buttons, insertion box and others
        self.create_table_button = None
        self.get_value = None  # Needs for radio buttons
        self.include_friday = None
        self.table_order_default = None
        self.table_order_sorted = None
        self.table_name_insertion_box = None

        # A label which will keep updating once user choose a data file
        self.button_text = tk.StringVar()
        self.button_text.set("File(s) Selected: ")
        self.create_files_names = Button(self.second_step_window, border=0, textvariable=self.button_text, command=self.changes_window_open,
                                         foreground="gray", font=("Arial", 11, "bold"))

        # Stores all errors found from receiverr.py
        self.error_data_list = []

        # Deletes previous files
        shutil.rmtree('copy_folder', ignore_errors=True)
        shutil.rmtree('__excel_files', ignore_errors=True)

        self.selection_step_window()

    def submit_tickcet_form(self):
        """Opens a Google Form to collect any reports or requests"""

        webbrowser.open("https://goo.gl/forms/wNkzjymOQ7wiNavf1")

    def open_pdf(self):
        """Instructions on how to use this program. Extremely useful"""

        return os.startfile("Uni-Schedule Instructions.pdf")

    def main_text_interface(self, button_frame, x=52, include_instructions=True):
        """Having repeated text such as main text, description,
        issue report icon, and information icon"""

        # What is the program main build for
        main_text = ttk.Label(button_frame,
                              text="Schedule Builder",
                              foreground="green",
                              font=('Courier', 20, 'bold'))
        main_text.grid(sticky='W',
                       column=0,
                       row=0,
                       rowspan=2,
                       padx=x,
                       pady=10)

        # What it does/creates
        main_text_functionality = ttk.Label(button_frame, text="Room table, master table and others... ",
                                            foreground="gray",
                                            font=("Courier", 10, 'bold'))
        main_text_functionality.grid(sticky="SW",
                                     column=0,
                                     row=1,
                                     padx=x)

        # Creating buttons and setting location for them
        problem_button = Button(button_frame,
                                border='0',
                                text="Report a problem or request",
                                command=self.submit_tickcet_form,
                                foreground="blue",
                                font=('Arial', 11, 'underline'))
        problem_button.grid(sticky="NE",
                            column=2,
                            row=1,
                            pady=2,
                            padx=3)

        if include_instructions is True:

            info_button = Button(button_frame,
                                 border='0',
                                 text="Instructions/Information",
                                 foreground="blue",
                                 command=self.open_pdf,
                                 font=('Arial', 11, 'underline'))
            info_button.grid(sticky="SE",
                             column=2,
                             row=2,
                             padx=3,
                             pady=0)
            info_button.config(image=self.info_image,
                               compound=RIGHT)

        problem_button.config(image=self.fix_image,
                              compound=RIGHT)

    def interface_window_remover(self):
        """Removes a window once a user goes to a next step or previous step"""

        if self.second_step_window:
            self.second_step_window.grid_remove()

        if self.first_step_window:
            self.first_step_window.grid_remove()

        if self.settings_window:
            self.settings_window.grid_remove()

        if self.creating_step_window:
            self.creating_step_window.grid_remove()

        if self.notification_window:
            self.notification_window.grid_remove()

    def excel_template_instructions(self):

        self.interface_window_remover()

        button_frame = self.first_step_window = Frame(self)
        button_frame.grid()
        self.create_files_names.place_forget()

        self.main_text_interface(button_frame, x=10)

        # Empty row for design purposes.
        empty_row = ttk.Label(button_frame,
                              text=" ")
        empty_row.grid(columnspan=3,
                       row=3)

        open_excel_template = Button(button_frame,
                                     relief="groove",
                                     bg='#c5eb93',
                                     border='4',
                                     text="Open a template",
                                     command=self.open_excel_template,
                                     foreground="green",
                                     font=('Arial', 20, 'bold'))

        open_excel_template.grid(sticky="S",
                                 columnspan=3,
                                 row=4,
                                 pady=10)

        excel_template_description = tk.Label(button_frame,
                                              text='Use the following template to assign the courses. This template will help to keep all the \n'
                                                   'information in the same format. Each department can have one template or more. If you     \n'
                                                   'want to change anything in the  template - please click "Report a problem or request"       ',
                                              foreground="gray",
                                              padx=17,
                                              font=("Arial", 11, 'bold'))
        excel_template_description.grid(row=6,
                                        column=0,
                                        columnspan=3,
                                        sticky='W')

        next_step_button = Button(button_frame,
                                  relief="groove",
                                  bg='#c5eb93',
                                  border='4',
                                  text="Move to next step",
                                  command=self.selection_step_window,
                                  foreground="green",
                                  font=('Arial', 16, 'bold'))
        next_step_button.grid(row=6,
                              sticky="NE",
                              rowspan=9,
                              column=2,
                              padx=3,
                              pady=88)

    def open_excel_template(self):
        """Includes an excel course schedule template which been agreed as a good template to use"""
        os.startfile('__Excel_Template.xltx')

    def select_excel_files(self):
        """Once a user selects the file - it will hold in the list."""

        self.file_name = Frame(self).filename = filedialog.askopenfilenames(initialdir="/",
                                                                            title="Select Excel file",
                                                                            filetypes=(("excel files", "*.xlsx"),
                                                                                       ("all files", "*.*")))
        if not self.file_name:
            pass
        else:
            for filesAmount in range(len(self.file_name)):
                self.files_show_directory.append(self.file_name[filesAmount])
                self.show_excel_files()

    def show_excel_files(self):
        """Shows to the user which files has been chosen"""
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

        elif len(self.files_show_names) >= 2:
            self.files_string = ("File(s) Selected: " + ", ".join(self.files_show_names))

            while len(self.files_string) > 83:
                self.files_string = self.files_string[:-1]

            if len(self.files_string) >= 83:
                self.files_string = self.files_string + "...\n"

            self.update_button_text(self.files_string)

    def update_button_text(self, text):
        self.button_text.set(text)

    def changes_window_open(self):
        pass

    def selection_step_window(self):
        if self.first_step_window:
            self.first_step_window.grid_remove()

        if self.settings_window:
            self.settings_window.grid_remove()

        button_frame = self.second_step_window = Frame(self)
        button_frame.grid()

        self.main_text_interface(button_frame, x=10)

        """
        back_button = Button(button_frame,
                             border='0',
                             image=self.back_image,
                             command=self.excel_template_instructions)
        back_button.grid(sticky='WN',
                         column=0,
                         row=1,
                         rowspan=2,
                         pady=7,
                         padx=3)
        """

        # Empty row for design purpose
        empty_row = ttk.Label(button_frame,
                              text=" ")
        empty_row.grid(columnspan=3,
                       row=3)

        select_all_files = Button(button_frame,
                                  relief="groove",
                                  bg='#c5eb93',
                                  border='4',
                                  text="Select all Excel files to continue",
                                  command=self.select_excel_files,
                                  foreground="green",
                                  font=('Arial', 18, 'bold'))
        select_all_files.place(x=126, y=120)
        self.create_files_names.place(x=8, y=207)  # Files selected category

        button_select_description = tk.Label(button_frame,
                                             text='Select an excel file/files which you would '
                                                  'like to make a table from',
                                             foreground="gray",
                                             font=("Arial", 10, 'bold'))
        button_select_description.place(x=105, y=178)

        difference_explanation_text = tk.Button(button_frame,
                                                border=0,
                                                text='Change/View/Delete file(s)',
                                                foreground="gray",
                                                font=("Arial", 10, "bold", 'underline'))
        difference_explanation_text.place(x=8, y=246)  # Files selected category

        self.create_table_button = Button(button_frame,
                                          relief="groove",
                                          bg='#c5eb93',
                                          border='4',
                                          text="Create an Excel table",
                                          command=self.table_setting_window,
                                          foreground="green",
                                          font=('Arial', 16, 'bold'))
        self.create_table_button.grid(sticky='e',
                                      column=1,
                                      columnspan=2,
                                      row=8,
                                      pady=167,
                                      padx=43)

        if not self.files_show_directory:

            self.create_table_button.configure(bg="#d9dad9",
                                               relief=SUNKEN,
                                               border='1',
                                               state="disabled")
        else:
            pass

    def table_setting_window(self):
        """Gives the ability to provide additional changes to the table if the user wants to."""

        if self.second_step_window:
            self.second_step_window.grid_remove()
            self.create_files_names.place_forget()

        button_frame = self.settings_window = Frame(self)
        button_frame.grid()

        self.main_text_interface(button_frame)

        back_button = Button(button_frame,
                             border='0',
                             image=self.back_image,
                             command=self.selection_step_window)
        back_button.grid(sticky='WN',
                         column=0,
                         row=1,
                         rowspan=2,
                         pady=7,
                         padx=3)

        empty_row = ttk.Label(button_frame,
                              text=" ")
        empty_row.grid(columnspan=3,
                       row=3)

        table_text = ttk.Label(button_frame,
                               text="Which type of order would you prefer?   ",
                               foreground="green",
                               font=('Arial', 13, 'bold'))
        table_text.grid(sticky='WN',
                        column=0,
                        row=2,
                        rowspan=3,
                        padx=13,
                        pady=20)

        # Select table type
        self.get_value = tk.IntVar()

        self.include_friday = tk.IntVar()

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

        self.table_order_sorted = Radiobutton(button_frame,
                                              text="Days in order",
                                              font=('Arial', 11),
                                              variable=self.get_value,
                                              command=self.user_table_choice,
                                              value=2)
        self.table_order_sorted.grid(column=0,
                                     row=3,
                                     sticky='w',
                                     padx=148)

        table_image_default = Button(button_frame,
                                     border='0',
                                     image=self.out_order_image,
                                     command=self.out_order_select)
        table_image_default.grid(column=0,
                                 row=4,
                                 sticky='w',
                                 padx=11)

        table_image_sorted = Button(button_frame,
                                    border='0',
                                    image=self.in_order_image,
                                    command=self.in_order_select)
        table_image_sorted.grid(column=0,
                                row=4,
                                sticky='w',
                                padx=150)

        # Table name
        get_table_name = ttk.Label(button_frame,
                                   text="Name the table",
                                   foreground="green",
                                   font=('Courier', 18, 'bold'))

        get_table_name.grid(column=0,
                            columnspan=3,
                            row=3,
                            sticky='ES',
                            padx=85,
                            pady=10)

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

        next_step_button = Button(button_frame,
                                  relief="groove",
                                  bg='#c5eb93',
                                  border='4',
                                  text="Create table",
                                  command=self.user_friday_choice,
                                  foreground="green",
                                  font=('Arial', 16, 'bold'))

        next_step_button.grid(column=0,
                              columnspan=3,
                              row=5,
                              rowspan=5,
                              sticky='EN',
                              pady=0,
                              padx=37)

        self.table_name_insertion_box.insert(END, "    Type name here...")
        self.table_name_insertion_box.bind("<1>", self.name_of_table)

        self.table_name_insertion_box.configure(font=('Courier', 12, 'italic'),
                                                foreground="gray")

        semesters_options = ["Fall",
                             "Spring",
                             "Summer"]

        self.table_name_insertion_box.bind("<Leave>", self.return_name)

        year_options = []

        for i in range(4):
            year_options.append(datetime.date.today().year + (i-1))

        variable_semesters = StringVar(button_frame)
        variable_years = StringVar(button_frame)
        today_year = datetime.datetime.now()
        variable_semesters.set(semesters_options[0])

        # Set the current year automatically
        for i in year_options:
            if today_year.year == i:
                variable_years.set(today_year.year)
            else:
                pass

        semester_text = ttk.Label(button_frame,
                                  text="Select the semester \nand year: ",
                                  foreground="green",
                                  font=('Courier', 15, 'bold'))

        semester_text.grid(column=0,
                           columnspan=3,
                           row=4,
                           sticky="e",
                           padx=83)

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

    def name_of_table(self, event):
        inp = self.table_name_insertion_box.get("1.0", END)
        if inp[:21] == "    Type name here...":
            self.table_name_insertion_box.delete("1.0", END)
            self.table_name_insertion_box.insert(END, " ")
            self.table_name_insertion_box.configure(font=('Courier', 12, 'bold'),
                                                    foreground="black")

    def user_friday_choice(self):
        self.table_friday_include = self.include_friday.get()
        self.program_loading_window()

    def user_table_choice(self):
        self.table_settings_type = self.get_value.get()

    def in_order_select(self):
        self.table_order_sorted.select()
        self.table_settings_type = 2

    def out_order_select(self):
        self.table_order_default.select()
        self.table_settings_type = 1

    def return_year(self, year_value):
        self.table_settings_year = year_value

    def return_semester(self, semester_value):
        self.table_settings_semester = semester_value

    def return_name(self, event):
        self.table_settings_name = self.table_name_insertion_box.get("1.0", END)

    def open_master_table(self):
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

    def program_loading_window(self):
        if self.first_step_window:
            self.first_step_window.grid_remove()
        else:
            pass
        if self.settings_window:
            self.settings_window.grid_remove()
        else:
            pass

        button_frame = self.creating_step_window = Frame(self)
        button_frame.grid()

        global switch
        switch = False

        # Needs additional work.
        def processor():
            global switch

            wait_text = StringVar()
            while not switch:
                wait_label = tk.Label(button_frame, textvariable=wait_text,
                                      foreground="green",
                                      font=('Courier', 20, 'bold'))
                wait_label.grid(column=1, row=2, rowspan=2, padx=10, pady=10)

                wait_text.set("\r \n \n \n  Please Wait...")
                wait_label.configure(textvariable=wait_text)

                sys.stdout.flush()
                time.sleep(0.1)

        processor_threading = threading.Thread(target=processor, name="processor thread")
        processor_threading.start()
        button_frame.update()

        self.error_data_list = receiver.DataProcessor(self.files_show_directory, self.table_settings_name,
                                                      self.table_settings_semester, self.table_settings_year,
                                                      self.table_settings_type,
                                                      self.table_friday_include).get_excel_errors()
        switch = True
        self.user_result_window()

    def user_result_window(self):
        if self.creating_step_window:
            self.creating_step_window.grid_remove()

        def clear_error_data(error_data):
            clear_data_list = []
            clear_data_dict = {}
            for i in range(len(error_data)):
                if error_data[i-1].get("Color") == 'FF687B':
                    try:
                        if error_data[i-1].get("Comment") == error_data[i].get("Comment"):
                            pass
                        else:
                            clear_data_dict['Message'] = error_data[i-1].get("Comment")
                            clear_data_list.append(clear_data_dict.copy())
                    except IndexError:
                        pass
            return clear_data_list
        clear_error_list = clear_error_data(self.error_data_list)
        if len(clear_error_list) != 0:

            button_frame = self.notification_window = Frame(self)
            button_frame.grid()

            self.main_text_interface(button_frame, x=10)

            total_number_mistakes = str(int(len(clear_error_list) / 2))
            if total_number_mistakes == '0':
                total_number_mistakes = '1'
            total_number_mistakes = 'Possible mistakes: ' + total_number_mistakes
            red_title = ttk.Label(button_frame,
                                  text=total_number_mistakes,
                                  foreground="gray",
                                  font=('Courier', 14, 'bold'))
            red_title.grid(sticky='w',
                           column=0,
                           columnspan=4,
                           row=2,
                           padx=10)

            # Creating a scroll window of errors
            def scroll_error_messages():
                for i in range(len(clear_error_list)):
                    conflict_row_message = str(clear_error_list[i].get("Message"))
                    Label(frame, text=conflict_row_message, background="#ee8282").grid(row=i, column=0)

            def show_all_messages(event):
                canvas.configure(scrollregion=canvas.bbox("all"), width=600, height=60)

            error_message_frame = Frame(button_frame, relief=GROOVE, width=600, height=150, bd=1)
            error_message_frame.place(x=22, y=110)

            canvas = Canvas(error_message_frame)
            frame = Frame(canvas)

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

            user_excel_copies = Button(button_frame,
                                       border='0',
                                       image=self.excel_copy_fie,
                                       command=self.open_excel_copies)
            user_excel_copies.grid(sticky="NW",
                                   column=0,
                                   row=3,
                                   rowspan=4,
                                   padx=20,
                                   pady=90)

            excel_copies_text = Button(button_frame,
                                       border='0',
                                       text="Open excel copies",
                                       command=self.open_excel_copies,
                                       foreground="gray",
                                       font=('Arial', 11, 'bold'))
            excel_copies_text.grid(sticky="W",
                                   column=0,
                                   row=3,
                                   rowspan=4,
                                   padx=8,
                                   pady=193)

            open_main_excel = Button(button_frame,
                                     border='0',
                                     image=self.excel_main_file,
                                     command=self.open_master_table)
            open_main_excel.grid(sticky="NE",
                                 column=0,
                                 columnspan=2,
                                 row=3,
                                 rowspan=4,
                                 padx=10,
                                 pady=90)

            main_excel_text = Button(button_frame,
                                     border='0',
                                     text="Open master table",
                                     command=self.open_master_table,
                                     foreground="gray",
                                     font=('Arial', 11, 'bold'))
            main_excel_text.grid(sticky="WE",
                                 column=0,
                                 columnspan=3,
                                 row=3,
                                 rowspan=4,
                                 padx=160,
                                 pady=0)

            exit_button = Button(button_frame,
                                 border='0',
                                 image=self.exit_file,
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
            no_errors_message = ttk.Label(button_frame,
                                          text=instructions_message,
                                          foreground="green",
                                          font=('Arial', 24, 'bold'))
            no_errors_message.grid(column=0,
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
                                          command=self.submit_tickcet_form,
                                          foreground="blue",
                                          font=('Arial', 11, 'underline'))
            user_feedback_button.grid(column=0,
                                      columnspan=3,
                                      row=3,
                                      sticky="EW",
                                      pady=60)


def create_interface(argv):

    root = Tk()
    root.title("Uni-Table Maker")
    root.geometry("659x337")
    UserInterface(root)
    root.mainloop()


if __name__ == "__main__":
    create_interface(sys.argv)

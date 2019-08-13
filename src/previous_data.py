import datetime
import os
import string
import urllib.parse

import openpyxl
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.styles.borders import Border, Side
from openpyxl.worksheet.pagebreak import Break

from src.previous_semesters import ReceiveSemesters


class PreviousCourses:
    def __init__(self, department, semester, year, url_smemester=None, url_department=None):
        self.department = department
        self.year = year
        self.semester = semester
        self.url_department = url_department
        self.url_semester = url_smemester

        self.headers = {
            'Content-Type': "application/x-www-form-urlencoded",
            'Origin': "https://secure2.mnsu.edu",
            'Referer': "https://secure2.mnsu.edu/courses/Default.asp",
            'cache-control': "no-cache",
            'Postman-Token': "6f1fa71c-c6fa-4fc4-b7df-e3cefe723179"
        }
        self.payload = None
        self.course_list = []
        self.main_class_controller()

    def main_class_controller(self):
        def transfer_department_name(department_abbreviation):
            """Transfers to the full department name"""
            return {
                'ACCT': 'Accounting',
                'BLAW': 'Business Law',
                'FINA': 'Finance',
                'IBUS': 'International Business',
                'MGMT': 'Management',
                'MRKT': 'Marketing',
                'MACC': 'Master in Accounting'
            }.get(department_abbreviation, default='Master of Business Administration')

        def get_payload_encode(encode_params, url, year, semester, department):
            """Gets parameters for semester and subject"""

            page_link = url
            page_response = requests.get(page_link)
            soup = BeautifulSoup(page_response.content, "html.parser")

            semester_option = semester + str(year)
            semester_option = semester_option.upper()
            department_option = (department.replace(" ", "")).upper()

            for option in soup.find_all('option'):
                search_option = (option.text.replace(" ", "")).upper()
                if semester_option == search_option:
                    encode_params['semester'] = option['value']
                if search_option[0:len(department_option)] == department_option:
                    encode_params['subject'] = option['value']

            return encode_params, semester_option

        def transfer_params(parse_params):
            """Urlparse"""
            parse_params = urllib.parse.urlencode(parse_params)

            params_list = [parse_params]
            return params_list

        full_department_name = transfer_department_name(department_abbreviation=self.department)
        params = {
            'semester': None,
            'campus': '1,2,3,4,5,6,7,9,A,B,C,I,L,M,N,P,Q,R,S,T,W,U,V,X,Y,Z',
            'startTime': '0600',
            'endTime': '2359',
            'days': 'ALL',
            'All': 'All Sections',
            'subject': None,
            'undefined': ''
        }
        # if not None, skips one function to increase performance
        if self.url_department is not None:
            params['subject'] = self.url_department
            params['semester'] = self.url_semester
            self.payload = transfer_params(params)
        else:
            web_params, sem_option = get_payload_encode(params, ReceiveSemesters.COURSES_URL, self.year, self.semester, full_department_name)
            self.payload = transfer_params(web_params)

        response = requests.request("POST", ReceiveSemesters.COURSES_URL, data=self.payload[0], headers=self.headers)
        self.get_data(response)
        CreateStandardTable(self.course_list, full_department_name, self.semester, str(self.year), self.department)

    def get_data(self, web_response):
        soup = BeautifulSoup(web_response.text, 'html.parser')

        for table_data in soup.find_all("tr"):
            course_data_part_one = []
            course_data_part_two = []

            course_title_raw = table_data.find(color="#ffffff")
            if course_title_raw is not None:
                for course_title in course_title_raw("b"):
                    title_text = course_title.get_text()
                    course_data_part_one.append(title_text)
            if (table_data["bgcolor"] == "#E1E1CC") or (table_data["bgcolor"] == "#FFFFFF"):
                for course_data in table_data("td"):
                    course_data_text = course_data.get_text()
                    course_data_part_two.append(course_data_text)

            if course_data_part_one:
                self.course_list.append(course_data_part_one)
            if course_data_part_two:
                if len(course_data_part_two[0]) == 6:
                    self.course_list.append(course_data_part_two)


class CreateStandardTable:
    def __init__(self, raw_data, departament_full, semester, year, department_abbreviation):
        self.raw_data = raw_data
        self.departament = departament_full
        self.semester = semester
        self.year = year
        self.abb_department = department_abbreviation

        self.workbook = None
        self.sheet = None

        self.main_program_controller()

    def main_program_controller(self):
        self.create_excel_file()

        self.create_excel_sheet()

        self.department_heading()
        self.column_headings()
        self.set_data()
        self.adjust_cells_width()
        self.border_all_cells("A1")
        self.set_page_break()
        self.workbook.save('web_files\\' + self.abb_department.replace(" ", "_")[0:27] + "_" +
                           self.year + ".xlsx")

    def create_excel_file(self):
        def create_directory():
            """Creates directory for created excel file"""
            if not os.path.exists('web_files'):
                os.makedirs('web_files')
        create_directory()

        self.workbook = openpyxl.Workbook()

    def create_excel_sheet(self):
        file_date = datetime.datetime.today().strftime('%Y')

        self.sheet = self.workbook.get_sheet_by_name("Sheet")
        self.sheet.title = self.departament + "_" + file_date

    def department_heading(self):
        self.sheet["A1"] = self.departament.upper() + " " + self.semester.upper() + " " + self.year
        self.sheet.merge_cells("A1:M2")
        self.sheet["A1"].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        self.sheet["A1"].font = Font(sz=12, bold=True, italic=False)

        clr = PatternFill(start_color='EEEEEE', end_color='EEEEEE', fill_type='solid')
        self.sheet["A1"].fill = clr

    def column_headings(self):
        col_headings = ["Course", "Number", "Section", "Credits", "Title of Course", "Room", "Days", "Time",
                        "Enrollment", "Faculty", "Start Date", "End Date", "Notes"]
        for col in range(len(col_headings)):
            coordinate = ''.join(string.ascii_uppercase[col]) + "3"  # i.e "A3"
            self.sheet[coordinate] = col_headings[col]
            self.sheet[coordinate].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            self.sheet[coordinate].font = Font(sz=12, bold=False, italic=False)

    def set_data(self):
        index = 0
        start_row = 4

        def insert_cells(raw_data, sheet, c_index, c_start_row):
            raw_data = raw_data
            sheet = sheet
            c_index = c_index
            c_start_row = c_start_row
            if raw_data[c_index][0]:
                if raw_data[c_index][0].isdigit() is not True:
                    course = raw_data[c_index][0]
                    course_number = raw_data[c_index][1]
                    symbol_index = raw_data[c_index][3].find("(")
                    course_credits = str(raw_data[c_index][3][symbol_index+1:symbol_index+2])
                    course_name = str(raw_data[c_index][2]).strip()

                    while len(raw_data[c_index + 1][0]) == 6 and (raw_data[c_index + 1][0].isdigit()):
                        sheet["A"+str(c_start_row)] = course
                        sheet["B" + str(c_start_row)] = int(course_number)
                        sheet["C" + str(c_start_row)] = int(raw_data[c_index + 1][1][0:2])
                        sheet["D" + str(c_start_row)] = int(course_credits)
                        sheet["E" + str(c_start_row)] = course_name
                        if raw_data[c_index + 1][4][0:3] == "ARR":
                            sheet["F" + str(c_start_row)] = "ARR"
                        else:
                            sheet["F" + str(c_start_row)] = raw_data[c_index + 1][6]
                        sheet["G" + str(c_start_row)] = str(raw_data[c_index + 1][3]).strip()
                        if raw_data[c_index + 1][6].replace(" ", "") == "ONLINE":
                            sheet["F" + str(c_start_row)] = "ONLINE"
                            sheet["H" + str(c_start_row)] = "ONLINE"
                        else:
                            sheet["H" + str(c_start_row)] = raw_data[c_index + 1][4]
                        sheet["I" + str(c_start_row)] = int(raw_data[c_index + 1][9])
                        sheet["J" + str(c_start_row)] = raw_data[c_index + 1][7]

                        date = raw_data[c_index + 1][5].split("-")
                        date[0] = date[0].replace(" ", "")
                        date[0] = date[0][:-2] + '20' + date[0][-2:]
                        date[1] = date[1].replace(" ", "")
                        date[1] = date[1][:-2] + '20' + date[1][-2:]

                        sheet["K" + str(c_start_row)] = date[0]
                        sheet["L" + str(c_start_row)] = date[1]

                        for col in range(12):
                            coordinate = ''.join(string.ascii_uppercase[col]) + str(c_start_row)
                            self.sheet[coordinate].alignment = Alignment(horizontal='center', vertical='center',
                                                                         wrap_text=True)

                        c_index += 1
                        c_start_row += 1
                    else:
                        insert_cells(raw_data, sheet, c_index + 1, c_start_row)

        try:
            insert_cells(self.raw_data, self.sheet, index, start_row)
        except IndexError:
            pass

    def adjust_cells_width(self):
        """Adjusts all the cell width. It is really cool"""

        for column in self.sheet.columns:
            max_length = 0
            get_column = column[0].column
            if get_column is "A":
                self.sheet.column_dimensions["A"].width = 9
            elif get_column is "B":
                self.sheet.column_dimensions["A"].width = 7.9

            else:
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except TypeError:
                        pass
                # A formula for auto adjusted width
                adjusted_width = (max_length + 2) * 1.02

                # Limit width
                if adjusted_width > 25:
                    adjusted_width = 25

                self.sheet.column_dimensions[get_column].width = adjusted_width

    def border_all_cells(self, start_cell):
        """Borders all table"""
        # Gets table size
        excel_max_row = self.sheet.max_row
        excel_max_column = self.sheet.max_column

        # Style of a border
        thin_border = Border(left=Side(style='thin'),
                             right=Side(style='thin'),
                             top=Side(style='thin'),
                             bottom=Side(style='thin'))

        # Goes over each cell and applies border
        for column in range(excel_max_column):
            # Transfers column to an alphabetical format
            col_letter = ''.join(string.ascii_uppercase[column])
            for row in range(excel_max_row):
                row += 1
                self.sheet[col_letter + str(row)].border = thin_border

    def set_page_break(self):
        # 40 rows per page
        get_max_row = self.sheet.max_row
        get_max_column = self.sheet.max_column

        if get_max_row >= 32:
            while get_max_row >= 32:
                self.sheet.sheet_properties.pageSetUpPr.fitToPage = True
                openpyxl.worksheet.pagebreak.PageBreak.tagname = 'rowBreaks'
                page_break_row = Break((get_max_row + 1) - 31)
                self.sheet.page_breaks.append(page_break_row)

                openpyxl.worksheet.pagebreak.PageBreak.tagname = 'colBreaks'
                page_break_column = Break(get_max_column + 1)
                self.sheet.page_breaks.append(page_break_column)
                get_max_row -= 32
        else:
            self.sheet.sheet_properties.pageSetUpPr.fitToPage = True
            openpyxl.worksheet.pagebreak.PageBreak.tagname = 'rowBreaks'
            page_break_row = Break(get_max_row + 1)
            self.sheet.page_breaks.append(page_break_row)

            openpyxl.worksheet.pagebreak.PageBreak.tagname = 'colBreaks'
            page_break_column = Break(get_max_column + 1)
            self.sheet.page_breaks.append(page_break_column)
        # Landscape orientation
        self.sheet.page_setup.orientation = self.sheet.ORIENTATION_LANDSCAPE

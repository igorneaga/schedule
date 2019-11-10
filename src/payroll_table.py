from nameparser import HumanName
import requests
from bs4 import BeautifulSoup
import re
import datetime
import os
import csv
import openpyxl
import string

from pprint import pprint


class PayrollTable:
    def __init__(self, courses_dict):
        self.courses_dict = courses_dict
        self.allowed_departments = [
            'Accounting', 'Business Law', 'Business', 'Finance',
            'Marketing & International Business', 'MBA', 'MACC',
            'Management'
        ]

        self.professor_data = []
        self.faculty_dict = dict()
        self.payroll_dict_format()
        self.create_excel_table()

    def payroll_dict_format(self):
        def create_faculty_dict(courses_dict):
            result = dict()
            for course_len in range(len(courses_dict)):
                for dict_key, faculty_name in courses_dict[course_len].items():
                    if dict_key == "Faculty":
                        if faculty_name != "None":
                            result.setdefault(faculty_name, [])
                            result[faculty_name].append(courses_dict[course_len])
            return result

        def get_professor_type(f_dict):
            prof_data = []
            allowed_departments = [
                'Accounting', 'Business Law', 'Business', 'Finance',
                'International Business', 'MBA', 'MACC',
                'Management', 'Marketing'
            ]

            def fetch_page():
                url = f'http://www.mnsu.edu/find/people.php?givenname={first_name}&sn={last_name}&employeetype='
                response = requests.get(url)
                return response.text

            def fetch_professor_type():
                page = fetch_page()
                soup = BeautifulSoup(page, 'html.parser')
                for p in soup.find_all('p'):
                    match = re.search(
                        fr'Name:(?P<name>.+?)Department:(?P<department>.+?{".*?|.*?".join(allowed_departments)}.+?); '
                        fr'Title: (?P<title>.+?); Type: (?P<type>.+?)Address: (?P<address>.+?); '
                        fr'Phone: (?P<phone>.+?); Email: (?P<email>.+?);',
                        p.text, flags=re.I
                    )

                    if match:
                        # let's remove trailing whitespaces in extracted data
                        data = match.groupdict()
                        for key, value in data.items():
                            data[key] = value.strip()

                        return data

                return None

            for professor in f_dict:
                name = HumanName(professor)
                first_name = name.first
                last_name = name.last
                professor_type = fetch_professor_type()
                if professor_type is None:
                    first_name = ""
                    professor_type = fetch_professor_type()
                    if professor_type is not None:
                        prof_data.append(professor_type)
                else:
                    prof_data.append(professor_type)
            return prof_data

        faculty_dict = create_faculty_dict(self.courses_dict)
        self.professor_data = get_professor_type(faculty_dict)

        def get_csv_file(file):
            cost_center = dict()
            if os.path.isfile(file):
                with open(file) as csv_file:
                    read_csv_file = csv.DictReader(csv_file, delimiter=',')
                    for row in read_csv_file:
                        cost_center = dict(row)
            return cost_center

        csv_file_data = get_csv_file('department_cost.csv')
        final_dict = dict()
        for professor, courses in faculty_dict.items():
            first_name = HumanName(professor).first
            last_name = HumanName(professor).last
            current_professor = next(
                (i for i in self.professor_data
                 if re.search(fr'{first_name}.*?{last_name}', i['name'])),
                dict()
            )
            final_dict[professor] = dict()
            final_dict[professor]['professor'] = current_professor
            tmp_courses = []

            for course in courses:
                if course['Department'].lower() == 'business':
                    cost = csv_file_data.get(current_professor.get('department'))
                    if current_professor.get('department'):
                        course['Department'] = current_professor['department']
                else:
                    cost = csv_file_data.get(course['Department'])
                course['Cost'] = cost
                tmp_courses.append(course)
            final_dict[professor]['courses'] = tmp_courses

        self.faculty_dict = final_dict

    def print_department_info(self, department):
        filtered_dict_courses = dict()

        for professor, data in self.faculty_dict.items():
            filtered_courses = [course for course in data['courses'] if
                                re.search(fr'.*?{department}.*?', course['Department'])]
            if filtered_courses:
                filtered_dict_courses[professor] = dict()
                filtered_dict_courses[professor]['courses'] = filtered_courses
                filtered_dict_courses[professor]['professor'] = data.get("professor")

        return filtered_dict_courses


    def create_excel_table(self):
        for dep in self.allowed_departments:
            print(dep)
            department_courses = self.print_department_info(dep)
            print(department_courses)
            if department_courses:
                self.main_program_controller(dep)


    def main_program_controller(self, department):
        self.create_excel_file()

        self.create_excel_sheet("Payroll Table", True)
        self.dates_heading()
        self.instructions_text()
        self.department_text(department)
        self.merge_excel_cells()
        #self.insert_courses()

        self.workbook.save('__excel_files\\' + "Payroll_Test_" + department + ".xlsx")

    def create_excel_file(self):
        def create_directory():
            """Creates directory for created excel file"""
            if not os.path.exists('__excel_files'):
                os.makedirs('__excel_files')
        create_directory()

        self.workbook = openpyxl.Workbook()

    def create_excel_sheet(self, sheet_name, first_sheet=False):
        if first_sheet is True:
            self.sheet = self.workbook.get_sheet_by_name("Sheet")
            self.sheet.title = "Payroll Table"
        else:
            self.workbook.create_sheet(sheet_name)
            self.sheet = self.workbook.get_sheet_by_name(sheet_name)

    def dates_heading(self):
        self.sheet["A1"] = "Table created:"
        self.sheet["A2"] = datetime.datetime.today().strftime('%m/%d/%Y')

    def instructions_text(self):
        self.sheet["F1"] = "* Key Concept for Scheduling - *Max Credits for Fall=14"
        self.sheet["F2"] = "* Final Total Credits must be 24"
        self.sheet["F3"] = "* Max Credit for the Year 29 with Overload"

    def department_text(self, department):
        self.sheet["A4"] = department

    def merge_excel_cells(self):
        # Dates merge cells
        self.sheet.merge_cells("A1:B1")
        self.sheet.merge_cells("A2:B2")

        # Instructions merge cells
        self.sheet.merge_cells("F1:J1")
        self.sheet.merge_cells("F2:J2")
        self.sheet.merge_cells("F3:J3")




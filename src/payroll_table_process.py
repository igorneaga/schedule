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
    def __init__(self):
        self.courses_dict = [{'Course': 'ACCT 220-1', 'Room': 'MH 0211', 'Course_Days': ['Thursday', 'Thursday'], 'Row': 4, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_FALL_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 1, 'Course_Title': 'Acct Cycle Apps', 'Type': ['Classroom', 'Telepresence'], 'Enrollment': 28, 'Faculty': 'Pike, Byron                                  ', 'Semester': 'Fall', 'Start_Time': '02:00', 'End_Time': '03:15', 'Start_Date': datetime.datetime(2019, 8, 27, 0, 0), 'End_Date': datetime.datetime(2019, 10, 3, 0, 0), 'Department': 'Accounting'}, {'Course': 'ACCT 210-1', 'Credits': 3, 'Course_Title': 'Managerial Accounting', 'Faculty': 'Fingland, Sean                               ', 'Enrollment': 40, 'Course_Days': [], 'Start_Time': 'Online', 'End_Time': 'Online', 'Type': 'Hybrid', 'Row': 16, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_FALL_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Start_Date': datetime.datetime(2019, 8, 26, 0, 0), 'End_Date': datetime.datetime(2019, 12, 13, 0, 0), 'Department': 'Accounting'}, {'Course': 'ACCT 210-2', 'Room': 'AH 0220', 'Course_Days': ['Monday', 'Wednesday'], 'Row': 17, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_FALL_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 3, 'Course_Title': 'Managerial Accounting', 'Type': ['Classroom', 'Hybrid'], 'Enrollment': 30, 'Faculty': 'Brennan, Paul                                ', 'Semester': 'Fall', 'Start_Time': '12:30', 'End_Time': '01:45', 'Start_Date': datetime.datetime(2019, 8, 26, 0, 0), 'End_Date': datetime.datetime(2019, 12, 13, 0, 0), 'Department': 'Accounting'}, {'Course': 'ACCT 210-4', 'Room': 'MH 0102', 'Course_Days': ['Thursday', 'Thursday'], 'Row': 18, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_FALL_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 3, 'Course_Title': 'Managerial Accounting', 'Type': ['Classroom'], 'Enrollment': 65, 'Faculty': 'Fingland, Sean                               ', 'Semester': 'Fall', 'Start_Time': '09:30', 'End_Time': '10:45', 'Start_Date': datetime.datetime(2019, 8, 26, 0, 0), 'End_Date': datetime.datetime(2019, 12, 13, 0, 0), 'Department': 'Accounting'}, {'Course': 'ACCT 210-5', 'Room': 'MH 0209', 'Course_Days': ['Thursday', 'Thursday'], 'Row': 19, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_FALL_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 3, 'Course_Title': 'Managerial Accounting', 'Type': ['Classroom'], 'Enrollment': 36, 'Faculty': 'Fingland, Sean                               ', 'Semester': 'Fall', 'Start_Time': '11:00', 'End_Time': '12:15', 'Start_Date': datetime.datetime(2019, 8, 26, 0, 0), 'End_Date': datetime.datetime(2019, 12, 13, 0, 0), 'Department': 'Accounting'}, {'Course': 'ACCT 210-6', 'Credits': 3, 'Course_Title': 'Managerial Accounting', 'Faculty': 'Fingland, Sean                               ', 'Enrollment': 36, 'Course_Days': [], 'Start_Time': 'Online', 'End_Time': 'Online', 'Type': ['Online'], 'Row': 20, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_FALL_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Start_Date': datetime.datetime(2019, 8, 26, 0, 0), 'End_Date': datetime.datetime(2019, 12, 13, 0, 0), 'Department': 'Accounting'}, {'Course': 'ACCT 210-41', 'Room': 'AH 0220', 'Course_Days': ['Monday', 'Wednesday'], 'Row': 22, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_FALL_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 3, 'Course_Title': 'Managerial Accounting', 'Type': ['Classroom'], 'Enrollment': 4, 'Faculty': 'Brennan, Paul                                ', 'Semester': 'Fall', 'Start_Time': '12:30', 'End_Time': '01:45', 'Start_Date': datetime.datetime(2019, 8, 26, 0, 0), 'End_Date': datetime.datetime(2019, 12, 13, 0, 0), 'Department': 'Accounting'}, {'Course': 'ACCT 300-3', 'Room': 'TC 0082', 'Course_Days': ['Monday', 'Wednesday'], 'Row': 25, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_FALL_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 3, 'Course_Title': 'Inter Fin Acct I', 'Type': ['Classroom'], 'Enrollment': 44, 'Faculty': 'Pike, Byron                                  ', 'Semester': 'Fall', 'Start_Time': '11:00', 'End_Time': '12:15', 'Start_Date': datetime.datetime(2019, 8, 26, 0, 0), 'End_Date': datetime.datetime(2019, 12, 13, 0, 0), 'Department': 'Accounting'}, {'Course': 'ACCT 301-1', 'Room': 'MH 0211', 'Course_Days': ['Monday', 'Wednesday'], 'Row': 26, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_FALL_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 3, 'Course_Title': 'Inter Fin Acct II', 'Type': ['Classroom', 'Telepresence'], 'Enrollment': 25, 'Faculty': 'Habib, Abo-El-Yazeed                         ', 'Semester': 'Fall', 'Start_Time': '02:00', 'End_Time': '03:15', 'Start_Date': datetime.datetime(2019, 8, 26, 0, 0), 'End_Date': datetime.datetime(2019, 12, 13, 0, 0), 'Department': 'Accounting'}, {'Course': 'ACCT 301-2', 'Room': 'MH 0211', 'Course_Days': ['Monday', 'Wednesday'], 'Row': 27, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_FALL_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 3, 'Course_Title': 'Inter Fin Acct II', 'Type': ['Classroom', 'Telepresence'], 'Enrollment': 28, 'Faculty': 'Habib, Abo-El-Yazeed                         ', 'Semester': 'Fall', 'Start_Time': '09:30', 'End_Time': '10:45', 'Start_Date': datetime.datetime(2019, 8, 26, 0, 0), 'End_Date': datetime.datetime(2019, 12, 13, 0, 0), 'Department': 'Accounting'}, {'Course': 'ACCT 310-1', 'Room': 'AH 0320', 'Course_Days': ['Thursday', 'Thursday'], 'Row': 28, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_FALL_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 3, 'Course_Title': 'Management Acct I', 'Type': ['Classroom', 'Hybrid'], 'Enrollment': 30, 'Faculty': 'Rosacker, Kirsten                            ', 'Semester': 'Fall', 'Start_Time': '12:30', 'End_Time': '01:45', 'Start_Date': datetime.datetime(2019, 8, 26, 0, 0), 'End_Date': datetime.datetime(2019, 12, 13, 0, 0), 'Department': 'Accounting'}, {'Course': 'ACCT 320-1', 'Room': 'AH 0320', 'Course_Days': ['Monday', 'Wednesday'], 'Row': 30, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_FALL_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 3, 'Course_Title': 'Acct Information Systems', 'Type': ['Classroom', 'Hybrid'], 'Enrollment': 30, 'Faculty': 'Johnson, Steven                              ', 'Semester': 'Fall', 'Start_Time': '03:30', 'End_Time': '04:45', 'Start_Date': datetime.datetime(2019, 8, 26, 0, 0), 'End_Date': datetime.datetime(2019, 12, 13, 0, 0), 'Department': 'Accounting'}, {'Course': 'ACCT 320-2', 'Room': 'AH 0320', 'Course_Days': ['Monday', 'Wednesday'], 'Row': 31, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_FALL_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 3, 'Course_Title': 'Acct Information Systems', 'Type': ['Classroom'], 'Enrollment': 30, 'Faculty': 'Johnson, Steven                              ', 'Semester': 'Fall', 'Start_Time': '12:30', 'End_Time': '01:45', 'Start_Date': datetime.datetime(2019, 8, 26, 0, 0), 'End_Date': datetime.datetime(2019, 12, 13, 0, 0), 'Department': 'Accounting'}, {'Course': 'ACCT 410-1', 'Room': 'AH 0220', 'Course_Days': ['Thursday', 'Thursday'], 'Row': 37, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_FALL_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 3, 'Course_Title': 'Business Income Tax', 'Type': ['Classroom'], 'Enrollment': 36, 'Faculty': 'Rosacker, Kirsten                            ', 'Semester': 'Fall', 'Start_Time': '03:30', 'End_Time': '04:45', 'Start_Date': datetime.datetime(2019, 8, 26, 0, 0), 'End_Date': datetime.datetime(2019, 12, 13, 0, 0), 'Department': 'Accounting'}, {'Course': 'ACCT 499-16', 'Room': 'ARR', 'Course_Days': [], 'Row': 43, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_FALL_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 1, 'Course_Title': 'Individual Study of Acct', 'Type': 'Error', 'Enrollment': 1, 'Faculty': 'Rosacker, Kirsten                            ', 'Semester': 'Fall', 'Start_Time': '11:00', 'End_Time': '12:15', 'Start_Date': datetime.datetime(2019, 8, 26, 0, 0), 'End_Date': datetime.datetime(2019, 12, 13, 0, 0), 'Department': 'Accounting'}, {'Course': 'ACCT 677-14', 'Room': 'ARR', 'Course_Days': [], 'Row': 44, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_FALL_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 1, 'Course_Title': 'Individual Study', 'Type': 'Error', 'Enrollment': 0, 'Faculty': 'Habib, Abo-El-Yazeed                         ', 'Semester': 'Fall', 'Start_Time': '11:00', 'End_Time': '12:15', 'Start_Date': datetime.datetime(2019, 8, 26, 0, 0), 'End_Date': datetime.datetime(2019, 12, 13, 0, 0), 'Department': 'Accounting'}, {'Course': 'BUS 295-40', 'Room': 'MH 0102', 'Course_Days': ['Thursday'], 'Row': 48, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_FALL_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 2, 'Course_Title': 'Prep for Bus Careers', 'Type': ['Classroom'], 'Enrollment': 7, 'Faculty': 'Diegnau, Melissa                             ', 'Semester': 'Fall', 'Start_Time': '01:30', 'End_Time': '03:20', 'Start_Date': datetime.datetime(2019, 8, 26, 0, 0), 'End_Date': datetime.datetime(2019, 12, 13, 0, 0), 'Department': 'Business'}, {'Course': 'ACCT 200-6', 'Room': 'MH 0102', 'Course_Days': ['Monday', 'Wednesday'], 'Row': 11, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_SPRING_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 3, 'Course_Title': 'Financial Accounting', 'Type': ['Classroom', 'Hybrid'], 'Enrollment': 10, 'Faculty': 'DeRemer, Mark                                ', 'Semester': 'Fall', 'Start_Time': '03:30', 'End_Time': '04:45', 'Start_Date': datetime.datetime(2020, 1, 13, 0, 0), 'End_Date': datetime.datetime(2020, 5, 8, 0, 0), 'Department': 'Accounting'}, {'Course': 'ACCT 210-5', 'Room': 'MH 0103', 'Course_Days': ['Monday', 'Wednesday'], 'Row': 21, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_SPRING_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 3, 'Course_Title': 'Managerial Accounting', 'Type': ['Classroom'], 'Enrollment': 1, 'Faculty': 'Siagian, Ferdinand                           ', 'Semester': 'Fall', 'Start_Time': '02:00', 'End_Time': '03:15', 'Start_Date': datetime.datetime(2020, 1, 13, 0, 0), 'End_Date': datetime.datetime(2020, 5, 8, 0, 0), 'Department': 'Accounting'}, {'Course': 'ACCT 398-1', 'Room': 'ARR', 'Course_Days': [], 'Row': 32, 'File': 'C:/Users/Igor/Desktop/fixing/acct/ACCT_SPRING_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 0, 'Course_Title': 'CPT: Co-Op Experience', 'Type': 'Error', 'Enrollment': 0, 'Faculty': 'Johnson, Steven                              ', 'Semester': 'Fall', 'Start_Time': '12:30', 'End_Time': '01:45', 'Start_Date': datetime.datetime(2020, 1, 13, 0, 0), 'End_Date': datetime.datetime(2020, 5, 8, 0, 0), 'Department': 'Accounting'}]

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

                # Changing the semester

                # TODO: CHECK POSSIBILITY TO COMBINE WITH PREVIOUS TO DO

                # Getting semester
                course_year = course.get("Start_Date").year
                fall_start_date = datetime.datetime(year=course_year, month=8, day=26)
                spring_start_date = datetime.datetime(year=course_year, month=1, day=13)
                summer1_start_date = datetime.datetime(year=course_year, month=5, day=18)
                summer2_start_date = datetime.datetime(year=course_year, month=6, day=22)

                if (fall_start_date - datetime.timedelta(days=33)) <= (course.get("Start_Date")) <= \
                        (fall_start_date + datetime.timedelta(days=33)):
                    course["Semester"] = "Fall"

                elif (spring_start_date - datetime.timedelta(days=33)) <= (course.get("Start_Date")) <= \
                        (spring_start_date + datetime.timedelta(days=33)):
                    course["Semester"] = "Spring"

                elif (summer1_start_date - datetime.timedelta(days=21)) <= (course.get("Start_Date")) <= \
                        (summer1_start_date + datetime.timedelta(days=21)):
                    course["Semester"] = "Summer1"

                elif (summer2_start_date - datetime.timedelta(days=21)) <= (course.get("Start_Date")) <= \
                        (summer2_start_date + datetime.timedelta(days=21)):
                    course["Semester"] = "Summer2"

                else:
                    course["Semester"] = "None"

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


PayrollTable()
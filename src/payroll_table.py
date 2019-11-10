from nameparser import HumanName
import requests
from bs4 import BeautifulSoup
import re
import datetime
import os
import csv
import openpyxl
from pprint import pprint


class PayrollTable:
    def __init__(self):
        self.courses_dict = [{'Course': 'BUS 295-1', 'Room': 'AH 0101', 'Course_Days': ['Tuesday'], 'Row': 4, 'File': 'C:/Users/Igor/Desktop/fixing/ACCT_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 2, 'Course_Title': 'Prep for Bus Careers', 'Type': ['Classroom'], 'Enrollment': 6, 'Faculty': 'Diegnau, Melissa                             ', 'Semester': 'Fall', 'Start_Time': '09:00', 'End_Time': '10:50', 'Start_Date': datetime.datetime(2020, 1, 13, 0, 0), 'End_Date': datetime.datetime(2020, 5, 8, 0, 0), 'Department': 'Business'}, {'Course': 'BUS 295-2', 'Room': 'MH 0102', 'Course_Days': ['Tuesday'], 'Row': 5, 'File': 'C:/Users/Igor/Desktop/fixing/ACCT_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 2, 'Course_Title': 'Prep for Bus Careers', 'Type': ['Classroom'], 'Enrollment': 8, 'Faculty': 'Diegnau, Melissa                             ', 'Semester': 'Fall', 'Start_Time': '11:00', 'End_Time': '12:50', 'Start_Date': datetime.datetime(2020, 1, 13, 0, 0), 'End_Date': datetime.datetime(2020, 5, 8, 0, 0), 'Department': 'Business'}, {'Course': 'BUS 295-3', 'Credits': 2, 'Course_Title': 'Prep for Bus Careers', 'Faculty': 'Diegnau, Melissa                             ', 'Enrollment': 17, 'Course_Days': [], 'Start_Time': 'Online', 'End_Time': 'Online', 'Type': ['Online'], 'Row': 6, 'File': 'C:/Users/Igor/Desktop/fixing/ACCT_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Start_Date': datetime.datetime(2020, 1, 13, 0, 0), 'End_Date': datetime.datetime(2020, 5, 8, 0, 0), 'Department': 'Business'}, {'Course': 'BUS 397-1', 'Room': 'MH 0102', 'Course_Days': ['Monday', 'Wednesday'], 'Row': 7, 'File': 'C:/Users/Igor/Desktop/fixing/ACCT_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 3, 'Course_Title': 'IBE Practicum', 'Type': ['Classroom'], 'Enrollment': 2, 'Faculty': 'Bowyer, Shane                                ', 'Semester': 'Fall', 'Start_Time': '02:10', 'End_Time': '03:20', 'Start_Date': datetime.datetime(2020, 1, 13, 0, 0), 'End_Date': datetime.datetime(2020, 5, 8, 0, 0), 'Department': 'Business'}, {'Course': 'BUS 397-2', 'Room': 'PH 0114', 'Course_Days': ['Monday'], 'Row': 8, 'File': 'C:/Users/Igor/Desktop/fixing/ACCT_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 3, 'Course_Title': 'IBE Practicum', 'Type': ['Classroom'], 'Enrollment': 1, 'Faculty': 'Scott, Kristin                               ', 'Semester': 'Fall', 'Start_Time': '02:10', 'End_Time': '03:20', 'Start_Date': datetime.datetime(2020, 1, 13, 0, 0), 'End_Date': datetime.datetime(2020, 5, 8, 0, 0), 'Department': 'Business'}, {'Course': 'BUS 397-3', 'Room': 'HC 1700 B', 'Course_Days': ['Monday'], 'Row': 9, 'File': 'C:/Users/Igor/Desktop/fixing/ACCT_2020.xlsx', 'Sheet_Name': 'Accounting_2019', 'Credits': 3, 'Course_Title': 'IBE Practicum', 'Type': ['Classroom'], 'Enrollment': 1, 'Faculty': 'Severns, Roger                               ', 'Semester': 'Fall', 'Start_Time': '02:10', 'End_Time': '03:20', 'Start_Date': datetime.datetime(2020, 1, 13, 0, 0), 'End_Date': datetime.datetime(2020, 5, 8, 0, 0), 'Department': 'Business'}]
        self.allowed_departments = [
            'Accounting', 'Business Law', 'Business', 'Finance',
            'Marketing & International Business', 'MBA', 'MACC',
            'Management'
        ]

        self.professor_data = []
        self.faculty_dict = dict()
        self.payroll_dict_format()

    def payroll_dict_format(self):
        allowed_departments = self.allowed_departments

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
        for professor, courses in faculty_dict.items():
            first_name = HumanName(professor).first
            last_name = HumanName(professor).last
            for course in courses:
                current_professor = next(
                    (i for i in self.professor_data
                     if re.search(fr'{first_name}.*?{last_name}', i['name'])),
                    dict()
                )
                if course['Department'] == 'Business':
                    if (csv_file_data.get('BUS')).lower() == 'professor':
                        cost = csv_file_data.get(current_professor.get('department'))
                        print(cost)
                        if current_professor.get('department'):
                            course['Department'] = current_professor['department']
                else:
                    cost = csv_file_data.get(course['Department'])
                course['Cost'] = cost

        self.faculty_dict = faculty_dict

    def print_department_info(self, department):
        filtered_dict = dict()
        for professor, courses in self.faculty_dict.items():
            filtered_courses = [course for course in courses if
                                re.search(fr'.*?{department}.*?', course['Department'])]
            if filtered_courses:
                filtered_dict[professor] = filtered_courses
        #pprint(filtered_dict)
        return filtered_dict


PayrollTable()

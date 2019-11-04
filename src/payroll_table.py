from nameparser import HumanName
import requests
from bs4 import BeautifulSoup
import re


class PayrollTable:
    def __init__(self, courses_dict):
        self.courses_dict = courses_dict
        self.payroll_dict_format()

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
            adj_dict = dict()

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
                        fr'Name:.+?Department:.+?(?:{"|".join(allowed_departments)}).+?; Title: (.+?);.+',
                        p.text, flags=re.I
                    )
                    if match:
                        return match.group(1)
                return None

            for professor in f_dict:
                name = HumanName(professor)
                first_name = name.first
                last_name = name.last
                professor_type = fetch_professor_type()
                if professor_type is None:
                    first_name = ""
                    professor_type = fetch_professor_type()
                    if professor_type == "Adjunct":
                        adj_dict.setdefault(professor, f_dict.get(professor))

            def deleting_same_data():
                e_dict = f_dict.copy()
                for key_f in f_dict:
                    for key_a in adj_dict:
                        if f_dict.get(key_f) == adj_dict.get(key_a):
                            del e_dict[key_f]
                return e_dict

            employee_d = deleting_same_data()

            return employee_d, adj_dict
        faculty_dict = create_faculty_dict(self.courses_dict)
        employee_dict, adjunct_dict = get_professor_type(faculty_dict)
        print(employee_dict)
        print(adjunct_dict)





import sqlite3
from sqlite3 import Error
import requests
from bs4 import BeautifulSoup
import urllib.parse
import logging
import datetime
import os
from pprint import pprint


class PreviousCoursesDatabase:
    COURSES_URL = 'https://secure2.mnsu.edu/courses/selectform.asp'

    def __init__(self, first_run):
        logging.basicConfig(filename='PreviousCoursesDatabase.log', level=logging.DEBUG)

        self.database = r"PreviousCourses.db"

        self.request_headers = {
            'Content-Type': "application/x-www-form-urlencoded",
            'Origin': "https://secure2.mnsu.edu",
            'Referer': "https://secure2.mnsu.edu/courses/Default.asp",
            'cache-control': "no-cache",
            'Postman-Token': "6f1fa71c-c6fa-4fc4-b7df-e3cefe723179"
        }

        self.main_controller()
        os.remove("PreviousCourses.db")

    def main_controller(self):
        db_connection = self.create_database()
        self.get_previous_courses()
        self.get_semester_dates(db_connection)

    def create_database(self):
        sql_create_semester_table = """CREATE TABLE IF NOT EXISTS semester  (
                                            semester_id        	TEXT NOT NULL,
                                            semester_year      	TEXT,
                                            semester_name      	TEXT,
                                            semester_start_date	TEXT,
                                            semester_end_date  	TEXT,
                                            semester_term      	TEXT,
                                            PRIMARY KEY(semester_id)
                                        );"""

        sql_create_room_table = """CREATE TABLE IF NOT EXISTS rooms (
                                                    room_id TEXT NOT NULL,
                                                    room_type TEXT,
                                                    cob_room INTEGER,
                                                    PRIMARY KEY(room_id)
                                                );"""

        sql_create_course_table = """CREATE TABLE IF NOT EXISTS course (
                                                        course_identifier TEXT NOT NULL,
                                                        course_number TEXT NOT NULL,
                                                        course_title TEXT,
                                                        course_credits TEXT,
                                                        semester_id TEXT,
                                                        PRIMARY KEY(course_identifier,course_number)
                                                    );"""
        sql_create_faculty_table = """CREATE TABLE IF NOT EXISTS faculty  (
                                         faculty_id        	TEXT NOT NULL,
                                         faculty_first_name	TEXT,
                                         faculty_last_name 	TEXT,
                                         faculty_department	TEXT,
                                         faculty_title     	TEXT,
                                         faculty_type      	TEXT,
                                         faculty_address   	TEXT,
                                         faculty_phone    	TEXT,
                                         faculty_email     	TEXT,
                                        PRIMARY KEY(faculty_id)
                                    );"""

        sql_create_course_section_table = """CREATE TABLE IF NOT EXISTS course_sections  (
                                                 course_section_id	TEXT NOT NULL,
                                                 course_identifier	TEXT NOT NULL,
                                                 course_number    	TEXT NOT NULL,
                                                 grade_id         	TEXT,
                                                 day_name         	TEXT,
                                                 start_time       	TEXT,
                                                 end_time          	TEXT,
                                                 start_date       	TEXT,
                                                 end_date         	TEXT,
                                                 room_number      	TEXT,
                                                 faculty_id       	TEXT NOT NULL,
                                                 course_size        TEXT,
                                                 enroll           	TEXT,
                                                 status           	TEXT,
                                                 notes            	TEXT,
                                                 course_type      	TEXT,
                                                PRIMARY KEY(course_section_id),
                                                FOREIGN KEY(course_identifier, course_number)
                                                REFERENCES course(course_identifier, course_number),
                                                FOREIGN KEY(faculty_id)
                                                REFERENCES faculty(faculty_id)
                                            );"""
        sql_create_course_meeting_table = """CREATE TABLE IF NOT EXISTS course_meetings (
                                                course_meeting_id	TEXT NOT NULL,
                                                course_section_id	TEXT NOT NULL,
                                                day_of_the_week  	TEXT,
                                                start_time       	TEXT,
                                                end_time         	TEXT,
                                                room_number      	TEXT NOT NULL,
                                                PRIMARY KEY(course_meeting_id),
                                                FOREIGN KEY(course_section_id)
                                                REFERENCES course_sections(course_section_id),
                                                FOREIGN KEY(room_number)
                                                REFERENCES room(room_number)
                                            );"""

        sql_create_grade_method_table = """CREATE TABLE IF NOT EXISTS grade_method  (
                                                grade_id         	TEXT NOT NULL,
                                                grade_description	TEXT,
                                                PRIMARY KEY(grade_id)
                                            );"""

        sql_create_section_grade_method_table = """CREATE TABLE IF NOT EXISTS section_grade_method  (
                                                        section_grade_method_id	TEXT NOT NULL,
                                                        course_section_id      	TEXT NOT NULL,
                                                        grade_id               	TEXT NOT NULL,
                                                        PRIMARY KEY(section_grade_method_id),
                                                        FOREIGN KEY(course_section_id)
                                                        REFERENCES course_sections(course_section_id),
                                                        FOREIGN KEY(grade_id)
                                                        REFERENCES grade_method(grade_id)
                                                    );"""

        sql_create_semester_course_table = """CREATE TABLE IF NOT EXISTS semester_course  (
                                                semester_id       	TEXT NOT NULL,
                                                course_identifier 	TEXT NOT NULL,
                                                course_number     	TEXT NOT NULL,
                                                semester_duration 	TEXT,
                                                semester_course_id	TEXT NOT NULL,
                                                PRIMARY KEY(semester_course_id),
                                                FOREIGN KEY(course_identifier, course_number)
                                                REFERENCES course(course_identifier, course_number),
                                                FOREIGN KEY(semester_id)
                                                REFERENCES semester(semester_id)
                                                );"""

        def create_connection(db_file):
            database_connection = None
            try:
                database_connection = sqlite3.connect(db_file)
                return database_connection
            except Error as e:
                logging.error("create_connection error", exc_info=True)

            return database_connection

        def create_table(database_connection, sql_table):
            try:
                database_cursor = database_connection.cursor()
                database_cursor.execute(sql_table)
            except Error as e:
                logging.error(e)

        db_connection = create_connection(self.database)

        if db_connection is not None:
            create_table(db_connection, sql_create_semester_table)
            create_table(db_connection, sql_create_room_table)
            create_table(db_connection, sql_create_course_table)
            create_table(db_connection, sql_create_faculty_table)
            create_table(db_connection, sql_create_course_section_table)
            create_table(db_connection, sql_create_course_meeting_table)
            create_table(db_connection, sql_create_grade_method_table)
            create_table(db_connection, sql_create_section_grade_method_table)
            create_table(db_connection, sql_create_semester_course_table)
            logging.info("Successfully created a table in the database")
        else:
            logging.warning("Connection has not been established")
        return db_connection

    def get_semester_dates(self, db_conn):
        semesters = ["summer", "fall", "spring"]
        years = []

        today_year = (datetime.datetime.now()).year
        for i in range(2):
            years.append(str(today_year - 1 - i))
            years.append(str(today_year + i))
        years.sort()

        def fetch_page(semester, year):
            year = year
            semester = semester
            url = f'https://www.mnsu.edu/registrar/{semester}{year}cal.html'
            response = requests.get(url)
            if response.status_code == 404:
                return None
            else:
                return response.text

        def sql_insert_semester(conn, task):
            cur = conn.cursor()
            cur.execute("INSERT INTO semester(semester_id, semester_year, semester_name, semester_start_date, "
                        "semester_end_date, semester_term) VALUES(?, ?, ?, ?, ?, ?)", task)

        def fetch_semester_date(semester, year):
            b_first_session = ""
            b_second_session = ""
            e_first_session = ""
            e_second_session = ""

            page = fetch_page(semester, year)
            if page is None:
                return None, None, None, None
            else:
                soup = BeautifulSoup(page, 'html.parser')
                td = soup.find_all(['td'])
                if semester == "summer":
                    for d in range(len(td)):
                        if ("classes begin" in td[d].text.lower()) & ("first session" in td[d].text.lower()):
                            b_first_session = td[d - 1].text
                        elif ("classes begin" in td[d].text.lower()) & ("second session" in td[d].text.lower()):
                            b_second_session = td[d - 1].text
                        elif "closing" in td[d].text.lower():
                            e_second_session = td[d - 1].text
                        else:
                            pass
                    e_first_session = b_second_session
                else:
                    for d in range(len(td)):
                        if "classes begin" in td[d].text.lower():
                            b_first_session = td[d - 1].text
                        elif "graduation" in td[d].text.lower():
                            e_first_session = td[d - 1].text
                        else:
                            pass

                if (b_first_session != "") and (e_first_session != ""):
                    return b_first_session, b_second_session, e_first_session, e_second_session
                else:
                    return None, None, None, None

        for y in years:
            for s in semesters:
                begin_first_session, begin_second_session, end_first_session, end_second_session = fetch_semester_date(s, y[-2:])
                if (begin_first_session is None) or (end_first_session is None):
                    if begin_first_session is None:
                        if s == "fall":
                            begin_first_session = '08/22/' + y
                        elif s == "spring":
                            begin_first_session = '01/11/' + y
                        else:
                            begin_first_session = '05/20/' + y
                            begin_second_session = '06/24/' + y
                    if end_first_session is None:
                        if s == "fall":
                            end_first_session = '01/01/' + str((int(y) + 1))
                        elif s == 'spring':
                            end_second_session = '05/10/' + str((int(y) + 1))
                        else:
                            end_first_session = '6/21/' + y
                            end_second_session = '7/26/' + y

                if s != 'summer':
                    unique_semester_id = s + "1" + y
                    sql_task = (unique_semester_id, y, s, begin_first_session, end_first_session, 1)
                    sql_insert_semester(db_conn, sql_task)
                else:
                    unique_semester_id = s + "1" + y
                    sql_task = (unique_semester_id, y, s, begin_first_session, end_first_session, 1)
                    sql_insert_semester(db_conn, sql_task)
                    unique_semester_id = s + "2" + y
                    sql_task = (unique_semester_id, y, s, begin_second_session, end_second_session, 2)
                    sql_insert_semester(db_conn, sql_task)


    def get_previous_courses(self):
        courses_dict = dict()
        courses_dict_list = []

        def web_scrap_param(courses_url):
            param_d_list = []
            page_response = requests.get(courses_url, verify=False)
            web_courses_parser = BeautifulSoup(page_response.content, "html.parser")

            scraping_param_dict = dict()
            for option in web_courses_parser.find_all('option'):
                search_option = (option.text.replace(" ", "")).upper()
                if search_option[0:4] == "FALL":
                    scraping_param_dict[search_option] = option['value']
                if search_option[0:4] == "SPRI":
                    scraping_param_dict[search_option] = option['value']
            param_d_list.append(scraping_param_dict)
            return param_d_list

        def get_university_courses(courses_url, d_key, request_headers):
            def get_data(web_response):
                courses_list = []
                data_html_parser = BeautifulSoup(web_response.text, 'html.parser')

                for table_data in data_html_parser.find_all("tr"):
                    course_titles_list = []
                    course_data_list = []

                    course_title_raw = table_data.find(color="#ffffff")
                    if course_title_raw is not None:
                        for course_title in course_title_raw("b"):
                            title_text = course_title.get_text()
                            course_titles_list.append(title_text)
                    if (table_data["bgcolor"] == "#E1E1CC") or (table_data["bgcolor"] == "#FFFFFF"):
                        for course_data in table_data("td"):
                            course_data_text = course_data.get_text()
                            course_data_list.append(course_data_text)

                    if course_titles_list:
                        courses_list.append(course_titles_list)
                    if course_data_list:
                        if len(course_data_list[0]) == 6:
                            courses_list.append(course_data_list)
                return courses_list

            default_params = {
                                'semester': d_key,
                                'campus': '1,2,3,4,5,6,7,9,A,B,C,I,L,M,N,P,Q,R,S,T,W,U,V,X,Y,Z',
                                'startTime': '0600',
                                'endTime': '2359',
                                'days': 'ALL',
                                'All': 'All Sections',
                                'undefined': ''
                                }

            def transfer_params(parse_params):
                """Urlparse"""
                parse_params = urllib.parse.urlencode(parse_params)
                params_list = [parse_params]
                return params_list

            user_request_encode = transfer_params(default_params)

            response = requests.request("POST", courses_url, data=user_request_encode[0],
                                        headers=request_headers)
            c_list = get_data(response)
            return c_list

        param_dict_list = web_scrap_param(self.COURSES_URL)

        def sql_insert_semester_course(conn, task):
            cur = conn.cursor()
            cur.execute("INSERT INTO course(semester_id, course_identifier, course_number, semester_duration"
                        "VALUES(?, ?, ?, ?, ?, ?)", task)

        def sql_insert_course(conn, task):
            cur = conn.cursor()
            cur.execute("INSERT INTO course(course_identifier, course_number, course_title, course_credits, "
                        "semester_id) VALUES(?, ?, ?, ?, ?, ?)", task)

        def sql_insert_course_section(conn, task):
            cur = conn.cursor()
            cur.execute("INSERT INTO semester(course_section_id, course_identifier, course_number, grade_id, "
                        "day_name, start_time, end_time, start_date, end_date, room_number, faculty_id, course_size"
                        "enroll, status, notes, course_type) VALUES(?, ?, ?, ?, ?, ?)", task)

        for dictionary in param_dict_list:
            for key in dictionary:
                semester = key[0:-4]
                year = key[-4:]
                #courses = get_university_courses(self.COURSES_URL, dictionary.get(key), self.request_headers)
                courses = ['000794',
                      '08',
                      'OPT',
                      'MT HF  ',
                      '12:00 pm - 12:50 pm',
                      '01/13/20 - 05/08/20',
                      'WC 0351       ',
                      'Whitcomb, Austin                             ',
                      '35',
                      '27',
                      'Open',
                      '\xa0']
                courses_dict[semester] = {year: courses}
                courses_dict_list.append(courses_dict)
                pprint(courses_dict_list)
                #sql_task = (unique_semester_id, y, s, begin_first_session, end_first_session, 1)


PreviousCoursesDatabase(first_run=True)
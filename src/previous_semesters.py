import requests
from bs4 import BeautifulSoup


class ReceiveSemesters:
    COURSES_URL = 'https://secure2.mnsu.edu/courses/selectform.asp'

    def __init__(self):
        self.param_dict_list = []

        self.web_scrap_param()
        self.return_courses_semesters()

    def web_scrap_param(self):
        page_response = requests.get(self.COURSES_URL, verify=False)
        web_courses_parser = BeautifulSoup(page_response.content, "html.parser")

        # College of Business department at Minnesota State University, Mankato
        cob_departments = ["ACCOUNTING(ACCT)", "BUSINESSLAW(BLAW)", "FINANCE(FINA)", "MANAGEMENT(MGMT)",
                           "MARKETING(MRKT)", "INTERNATIONALBUSINESS(IBUS)",
                           "MASTEROFBUSINESSADMINISTRATION(MBA)", "MASTERINACCOUNTING(MACC)"]

        scraping_param_dict = dict()
        for option in web_courses_parser.find_all('option'):
            search_option = (option.text.replace(" ", "")).upper()
            if search_option[0:4] == "FALL":
                scraping_param_dict[search_option] = option['value']
            if search_option[0:4] == "SPRI":
                scraping_param_dict[search_option] = option['value']
            for department in cob_departments:
                if department == search_option:
                    scraping_param_dict[search_option] = option['value']
        self.param_dict_list.append(scraping_param_dict)

    def return_courses_semesters(self):
        return self.param_dict_list

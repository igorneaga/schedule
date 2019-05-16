import requests
from bs4 import BeautifulSoup


class ReceiveSemesters:
    def __init__(self):
        self.url = "https://secure2.mnsu.edu/courses/selectform.asp"

        self.list_of_dict = []

        self.web_scrap_param()
        self.return_courses_semesters()

    def web_scrap_param(self):
        page_link = self.url
        page_response = requests.get(page_link)
        soup = BeautifulSoup(page_response.content, "html.parser")

        cob_departments = ["ACCOUNTING(ACCT)", "BUSINESSLAW(BLAW)", "FINANCE(FINA)", "MANAGEMENT(MGMT)", "MARKETING(MRKT)", "INTERNATIONALBUSINESS(IBUS)",
                           "MASTEROFBUSINESSADMINISTRATION(MBA)", "MASTERINACCOUNTING(MACC)"]

        scraping_param_dict = dict()
        for option in soup.find_all('option'):
            search_option = (option.text.replace(" ", "")).upper()
            if search_option[0:4] == "FALL":
                scraping_param_dict[search_option] = option['value']
            if search_option[0:4] == "SPRI":
                scraping_param_dict[search_option] = option['value']
            for department in cob_departments:
                if department == search_option:
                    scraping_param_dict[search_option] = option['value']
        self.list_of_dict.append(scraping_param_dict)

    def return_courses_semesters(self):
        return self.list_of_dict




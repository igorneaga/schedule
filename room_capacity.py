import csv
from bs4 import BeautifulSoup
from bs4 import element
import requests
import os
from datetime import date, timedelta, datetime


class RoomCapacity:
    """Web scraping university website to gain data about classroom capacity"""
    def __init__(self):

        self.room_cap_url = 'https://mnsu.bookitadmin.minnstate.edu/BrowseForSpace.aspx'
        self.file = 'room_cap.csv'

    def check_file_exist(self):

        def get_room_capacity(url):
            page_link = url
            page_response = requests.get(page_link, verify=False)
            page_content = BeautifulSoup(page_response.content, "html.parser")
            list_of_rooms = []
            list_of_capacity = []
            room_cap_dict = dict()
            for div in page_content.findAll('div', {'class': 'v'}):
                building_html = div.find(("a", {"class": "sm t c"}))
                building_number = "".join([t for t in building_html.contents if type(t) == element.NavigableString])
                symbol_index = building_number.find("(")

                room_html = div.findAll("a", {"class": "sm rl f h"})
                capacity_html = div.findAll("div", {"class": "cl h"})

                for rooms in room_html:
                    list_of_rooms.append(building_number[symbol_index + 1:-1] + " " +
                                         (rooms.text[0:5].replace(" ", "")))
                for capacity in capacity_html:
                    list_of_capacity.append(capacity.text)

            for i in range(len(list_of_rooms)):
                room_cap_dict[list_of_rooms[i]] = list_of_capacity[i]
            return room_cap_dict

        if os.path.isfile(self.file):
            # If the file already exists
            with open(self.file) as csv_file:
                read_csv_file = csv.DictReader(csv_file, delimiter=',')
                for row in read_csv_file:
                    room_cap = dict(row)

                previous_date = datetime.strptime(room_cap.get("Date"), '%Y-%m-%d')
                if date.today() > previous_date.date() + timedelta(days=60):
                    # Will rewrite file if it older than 60 days
                    room_cap = get_room_capacity(self.room_cap_url)
                    room_cap["Date"] = date.today()
                    with open(self.file, 'w') as over_write_file:
                        write_file = csv.DictWriter(over_write_file, room_cap.keys())
                        write_file.writeheader()
                        write_file.writerow(room_cap)
                    return room_cap
                else:
                    return dict(room_cap)

        else:
            # Creates a CSV file
            room_cap = get_room_capacity(self.room_cap_url)
            room_cap["Date"] = date.today()
            with open(self.file, 'w') as new_file:
                write_file = csv.DictWriter(new_file, room_cap.keys())
                write_file.writeheader()
                write_file.writerow(room_cap)
            return room_cap

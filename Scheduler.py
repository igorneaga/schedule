import datetime
import io
import os
import queue as Queue
import shutil
import subprocess
import threading
import time
import tkinter as tk
import urllib.request
import urllib.request
import zipfile
from tkinter import *
from tkinter import ttk

import requests

API_GITHUB_UPDATE = "https://api.github.com/repos/igorneaga/schedule/releases/latest"
API_GITHUB_ASSETS = "https://api.github.com/repos/igorneaga/schedule/contents/src/assets"
MAIN_EXE_URL = "https://github.com/igorneaga/schedule/raw/master/src/UScheduler.exe"
MAIN_ZIP_URL = "https://github.com/igorneaga/schedule/archive/master.zip"


class UpdateInterface(Frame):
    def __init__(self, master):
        super().__init__(master)
        self.grid()

        self.install_update_window = None

        self.update_in_progress()

    def update_in_progress(self):
        """An updating window with a progress bar"""
        button_frame = self.install_update_window = Frame(self)
        button_frame.grid()

        main_text = tk.Label(button_frame,
                             text="Please wait...",
                             foreground="green",
                             font=('Arial', 21))
        # Placing coordinates
        main_text.grid(column=0,
                       row=0,
                       padx=175,
                       pady=30,
                       sticky="n")

        description_text = tk.Label(button_frame,
                                    text="Updating necessary files",
                                    foreground="gray",
                                    font=('Arial', 11, 'italic'))
        description_text.grid(column=0,
                              row=1,
                              padx=175,
                              pady=20,
                              sticky="n")
        self.start_process()

    def progress(self):
        """Progress bar settings and coordination"""
        self.progress_bar = ttk.Progressbar(
            self.master, orient="horizontal",
            length=200, mode="determinate"
        )
        self.progress_bar.place(x=157, y=85)

    def start_process(self):
        self.progress()
        self.progress_bar.start()
        self.queue = Queue.Queue()
        ThreadedTask(self.queue).start()
        self.master.after(100, self.process_queue)

    def exit_function(self):
        sys.exit()

    def process_queue(self):
        try:
            # Moves progress bar
            msg = self.queue.get(0)
            self.progress_bar.stop()
        except Queue.Empty:
            self.master.after(100, self.process_queue)
        if len(threading.enumerate()) == 1:
            # Stops if there is only one process going
            self.exit_function()


class ThreadedTask(threading.Thread):
    def __init__(self, queue):
        threading.Thread.__init__(self)
        self.queue = queue

    def run(self):
        def download_zip(directory, url):
            """Downloads & unzip if the assets folder is missing"""

            shutil.rmtree(directory + "\\src", ignore_errors=True)
            zip_response = requests.get(url, allow_redirects=True)
            zip_file = zipfile.ZipFile(io.BytesIO(zip_response.content))

            for file in zip_file.namelist():
                if file.startswith('schedule-master/src/'):
                    zip_file.extract(file)
            os.rename(f'{directory}\\schedule-master\\src', f'{directory}\\src')
            shutil.rmtree(directory + "\\schedule-master", ignore_errors=True)

        script_directory = os.path.dirname(os.path.abspath(__file__))
        # Assets
        page_response_assets = requests.get(API_GITHUB_ASSETS)
        github_assets_data = page_response_assets.json()
        # Version / Date
        page_response = requests.get(API_GITHUB_UPDATE)
        git_app_date = page_response.json().get("published_at")
        if os.path.isdir(script_directory + "\\src") is True:
            # Checks for main file
            if os.path.isfile(script_directory + "\\src\\UScheduler.exe") is False:
                urllib.request.urlretrieve(MAIN_EXE_URL, f'{script_directory}\\src\\UScheduler.exe')
            else:
                # Checks for new release
                file_date = os.path.getmtime(script_directory + "\\src\\UScheduler.exe")
                modification_time = time.strftime('%Y-%m-%d', time.localtime(file_date))
                if git_app_date[:10] > modification_time[:10]:
                    os.remove("src\\UScheduler.exe")
                    urllib.request.urlretrieve(MAIN_EXE_URL, f'{script_directory}\\src\\UScheduler.exe')

        if os.path.isdir('src\\assets') is False:
            download_zip(script_directory, MAIN_ZIP_URL)
        else:
            try:
                # Goes through assets files to find if any updated one exists.
                for github_assets in github_assets_data:
                    if github_assets.get("name") in os.listdir(f'{script_directory}\\src\\assets'):
                        pass
                    else:
                        urllib.request.urlretrieve(github_assets.get("download_url"), f'{script_directory}\\src\\assets')
                        current_date_time = datetime.datetime.now()
                        modified_time = time.mktime(current_date_time.timetuple())
                        os.utime(f'{script_directory}\\src\\UScheduler.exe', (modified_time, modified_time))

            # Running without administrator permission will need to go through a complicated way
            except PermissionError:
                download_zip(script_directory, MAIN_ZIP_URL)

        subprocess.Popen(f'{script_directory}\\src\\UScheduler.exe', close_fds=True)
        # Gives some time for the main file to launch
        time.sleep(7)
        self.queue.put("Task finished")


def create_interface(argv):
    root = Tk()
    root.title('Uni-Scheduler')
    root.geometry("520x175")

    # Gets both half the screen width/height and window width/height
    screen_middle_w = int((root.winfo_screenwidth() / 2) - (520 / 2))
    screen_middle_h = int(root.winfo_screenheight() / 2 - 175 / 2)

    # Positions the window in the center of the page.
    root.geometry("+{}+{}".format(screen_middle_w, screen_middle_h))

    try:
        root.iconbitmap('src\\assets\\unischeduler_icon.ico')
    except:
        pass

    UpdateInterface(root)
    root.mainloop()


if __name__ == "__main__":
    create_interface(sys.argv)

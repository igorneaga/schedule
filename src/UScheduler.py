from src import interface
from tkinter import *
import os
import time
import datetime


def create_interface(argv):
    cwd = os.getcwd()
    # Change file date for Auto-Updates purpose
    current_date_time = datetime.datetime.now()
    modified_time = time.mktime(current_date_time.timetuple())
    os.utime(cwd + '\\src\\UScheduler.exe', (modified_time, modified_time))

    root = Tk()
    root.title('Uni-Scheduler')
    root.geometry("659x337")
    root.iconbitmap(cwd + '\\src\\assets\\unischeduler_icon.ico')
    interface.UserInterface(root)
    root.mainloop()


if __name__ == "__main__":
    create_interface(sys.argv)


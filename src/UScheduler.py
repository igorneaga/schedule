from src import interface
from tkinter import *
import os


def create_interface(argv):

    root = Tk()
    root.title('Uni-Scheduler')
    root.geometry("659x337")
    cwd = os.getcwd()
    root.iconbitmap(cwd + '\\src\\assets\\unischeduler_icon.ico')
    interface.UserInterface(root)
    root.mainloop()


if __name__ == "__main__":
    create_interface(sys.argv)

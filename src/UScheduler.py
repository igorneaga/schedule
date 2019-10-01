import os
from tkinter import *

from src import interface


def create_interface(argv):
    cwd = os.getcwd()

    root = Tk()
    root.title('Uni-Scheduler')
    root.geometry("659x337")
    root.iconbitmap(f'{cwd}\\src\\assets\\unischeduler_icon.ico')
    interface.UserInterface(root)
    root.mainloop()


if __name__ == "__main__":
    create_interface(sys.argv)


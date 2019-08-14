from src import interface
from tkinter import *

version = 'v0.6'  # Current version


def create_interface(argv):

    root = Tk()
    root.title('Uni-Scheduler')
    root.geometry("659x337")
    root.iconbitmap('assets\\unischeduler_icon.ico')
    interface.UserInterface(root)
    root.mainloop()


if __name__ == "__main__":
    create_interface(sys.argv)
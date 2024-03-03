import gui7
from tkinter import *


class HelloPackage(gui7.HelloPackage):
    def __getattr__(self, name):
        return getattr(self.top, name)   # передать вызов настоящему виджету

if __name__ == '__main__': HelloPackage().mainloop()   # вызовет __getattr__!

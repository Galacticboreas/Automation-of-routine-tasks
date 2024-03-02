from tkinter import *             # импортировать клаccы виджетов Tk
from gui6 import Hello            # импортировать подкласс фрейма

class HelloContainer(Frame):
    def __init__(self, parent=None):
        Frame.__init__(self, parent)
        self.pack()
        self.makeWidgets()
    
    def makeWidgets(self):
        Hello(self).pack(side=RIGHT)   # прикрепить объект класса Hello к себе
        Button(self, text='Attach', command=self.quit).pack(side=LEFT)

if __name__ == '__main__': HelloContainer().mainloop()

"""
кнопка Quit, которая запрашивает подтверждение на завершение;
для повторного использования достаточно прикрепить экземпляр к другому
графическому интерфейсу и скомпоновать с желаемыми параметрами
"""

from tkinter import *                      # импортировать классы виджетов
from tkinter.messagebox import askokcancel # импортировать стандартный диалог

class Quitter(Frame):                      # подкласс графич. интерфейса
    def __init__(self, parent=None):       # метод конструктора
        Frame.__init__(self, parent)
        self.pack()
        widget = Button(self, text='Quit', command=self.quit)
        widget.pack(side=LEFT, expand=YES, fill=BOTH)
    
    def quit(self):
        ans = askokcancel('Verify exit', "Really quit?")
        if ans: Frame.quit(self)

if __name__ == '__main__': Quitter().mainloop()

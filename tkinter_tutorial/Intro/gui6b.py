from sys import exit
from tkinter import *            # импортировать классы виджетов Tk
from gui6 import Hello           # импортировать подкласс фрейма

parent = Frame(None)             # создать контейнерный виджет
parent.pack()
Hello(parent).pack(side=RIGHT)   # прикрепить виджет Hello, не запуская его

Button(parent, text='Attach', command=exit).pack(side=LEFT)
parent.mainloop()

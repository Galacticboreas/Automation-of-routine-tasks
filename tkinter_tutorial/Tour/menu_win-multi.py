from menu_win import makemenu   # повторно использовать функцию создания меню
from tkinter import *


root = Tk()
for i in range(3):         # три всплывающих окна
    win = Toplevel(root)
    makemenu(win)
    Label(win, bg='black', height=5, width=25).pack(expand=YES, fill=BOTH)

Button(root, text="Bye", command=root.quit).pack()
root.mainloop()

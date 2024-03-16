from menu_frm import makemenu   # нельзя использовать menu_win--корень=Frame
from tkinter import *


root = Tk()
for i in range(3):           # три меню, вложенные в контейнеры
    frm = Frame()
    mnu = makemenu(frm)
    mnu.config(bd=2, relief=RAISED)
    frm.pack(expand=YES, fill=BOTH)
    Label(frm, bg='black', height=5, width=25).pack(expand=YES, fill=BOTH)
    Button(root, text="Bye", command=root.quit).pack()
    root.mainloop()

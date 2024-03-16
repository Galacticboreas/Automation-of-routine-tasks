from menu_frm import makemenu   # здесь нельзя использовать menu_win-- одно окно
from tkinter import *           # но можно прикреплять меню на основе фреймов

root = Tk()
for i in range(2):            # 2 меню в одном окне
    mnu = makemenu(root)
    mnu.config(bd=2, relief=RAISED)
    Label(root, bg='black', height=5, width=25).pack(expand=YES, fill=BOTH)

Button(root, text="Bye", command=root.quit).pack()
root.mainloop()

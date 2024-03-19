"""
создает формы с применением методов pack и grid в отдельных фреймах в одном
и том же окне; методы grid и pack не могут одновременно использоваться в одном
родительском контейнере (например, в корневом окне), но могут использоваться
в разных фреймах в одном и том же окне;
"""

from tkinter import *
from grid2 import gridbox, packbox


root = Tk()

Label(root, text='Grid:').pack()
frm = Frame(root, bd=5, relief=RAISED)
frm.pack(padx=5, pady=5)
gridbox(frm)

Label(root, text='Pack:').pack()
frm = Frame(root, bd=5, relief=RAISED)
frm.pack(padx=5, pady=5)
packbox(frm)

Button(root, text='Quit', command=root.quit).pack()
mainloop()

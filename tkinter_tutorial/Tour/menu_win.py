# меню окна верхнего уровня в стиле Tk 8.0

from tkinter import *            # импортировать базовый набор виджетов
from tkinter.messagebox import * # импортировать стандартные диалоги


def notdone():
    showerror('Not implemented', 'Not yet available')


def makemenu(win):
    top = Menu(win)          # win = окно верхнего уровня
    win.config(menu=top)     # установить его параметр menu

    file = Menu(top)
    file.add_command(label='New...', command=notdone, underline=0)
    file.add_command(label='Open...', command=notdone, underline=0)
    file.add_command(label='Quit', command=win.quit, underline=0)
    top.add_cascade(label='File', menu=file, underline=0)

    edit = Menu(top, tearoff=False)
    edit.add_command(label='Cut', command=notdone, underline=0)
    edit.add_command(label='Paste', command=notdone, underline=0)
    edit.add_separator()
    top.add_cascade(label='Edit', menu=edit, underline=0)

    submenu = Menu(edit, tearoff=True)
    submenu.add_command(label='Spam', command=win.quit, underline=0)
    submenu.add_command(label='Eggs', command=notdone, underline=0)
    edit.add_cascade(label='Stuff', menu=submenu, underline=0)


if __name__ == '__main__':
    root = Tk()                      # или Toplevel()
    root.title('menu_win')           # информация для менеджера окно
    makemenu(root)                   # создать строку меню
    msg = Label(root, text='Window menu basics') # добавить что-нибудь ниже
    msg.pack(expand=YES, fill=BOTH)
    msg.config(relief=SUNKEN, width=40, height=7, bg='beige')
    root.mainloop()

# Меню на основе фреймов: пригодно для окон верхнего уровня и компонентов

from tkinter import *             # импортировать базовый набор виджетов
from tkinter.messagebox import *  # импортировать стандартные диалони


def notdone():
    showerror('Non implemented', 'Not yet available')


def makemenu(parent):
    menubar = Frame(parent)           # relief=RAISED, bd=2...
    menubar.pack(side=TOP, fill=X)

    fbutton = Menubutton(menubar, text='File', underline=0)
    fbutton.pack(side=LEFT)
    file = Menu(fbutton)
    file.add_command(label='New...', command=notdone, underline=0)
    file.add_command(label='Open...', command=notdone, underline=0)
    file.add_command(label='Quit', command=parent.quit, underline=0)
    fbutton.config(menu=file)

    ebutton = Menubutton(menubar, text='Edit', underline=0)
    ebutton.pack(side=LEFT)
    edit = Menu(ebutton, tearoff=False)
    edit.add_command(label='Cut', command=notdone, underline=0)
    edit.add_command(label='Paste', command=notdone, underline=0)
    edit.add_separator()
    ebutton.config(menu=edit)

    submenu = Menu(edit, tearoff=True)
    submenu.add_command(label='Spam', command=parent.quit, underline=0)
    submenu.add_command(label='Eggs', command=notdone, underline=0)
    edit.add_cascade(label='Stuff', menu=submenu, underline=0)
    return menubar

if __name__ == '__main__':
    root = Tk()                   # или TopLevel, или Fram
    root.title('menu_frm')        # информация для менеджера окон
    makemenu(root)                # создать строку меню
    msg = Label(root, text='Frame menu basics') # добавить что-нибудь ниже
    msg.pack(expand=YES, fill=BOTH)
    msg.config(relief=SUNKEN, width=40, height=7, bg='beige')
    root.mainloop()

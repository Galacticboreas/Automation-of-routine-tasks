# изменяет размеры изображений для кнопок на панели инструментов с помощью PIL

from tkinter import *                  # импортировать базовый набор виджетов
from tkinter.messagebox import *       # импортировать стандартные диалоги


class NewMenuDemo(Frame):              # расширенный фрейм

    def __init__(self, parent=None):   # прикрепляется к корневому окну?
        Frame.__init__(self, parent)   # вызвать метод суперкласса
        self.pack(expand=YES, fill=BOTH)
        self.createWidgets()           # прикрепить фреймы/виджеты
        self.master.title("Toolbars and Menus")   # для менеджера окна
        self.master.iconname("tkpython")          # текст метки при свертывании

    
    def createWidgets(self):
        self.makeMenuBar()
        self.makeToolBar()
        L = Label(self, text='Menu and Toolbar Demo')
        L.config(relief=SUNKEN, width=40, height=10, bg='white')
        L.pack(expand=YES, fill=BOTH)

    
    def makeToolBar(self, size=(40, 40)):
        imgdir = r'resources/'
        toolbar = Frame(self, cursor='hand2', relief=SUNKEN, bd=2)
        toolbar.pack(side=BOTTOM, fill=X)
        photos = ('ora-lp4e.gif', 'pythonPowered.gif',
                  'python_conf_ora.gif')
        self.toolPhotoObjs = []
        for file in photos:
            img = PhotoImage(file=imgdir + file)
            btn = Button(toolbar, image=img, command=self.greeting)
            btn.config(bd=5, relief=RIDGE)
            btn.config(width=size[0], height=size[1])
            btn.pack(side=LEFT)
            self.toolPhotoObjs.append(img)      # сохранить ссылку
        Button(toolbar, text='Quit', command=self.quit).pack(side=RIGHT, fill=Y)

    
    def makeMenuBar(self):
        self.menubar = Menu(self.master)
        self.master.config(menu=self.menubar)   # master=окно верхнего уровня
        self.fileMenu()
        self.editMenu()
        self.imageMenu()


    def fileMenu(self):
        pulldown = Menu(self.menubar)
        pulldown.add_command(label='Open...', command=self.notdone)
        pulldown.add_command(label='Quit', command=self.quit)
        self.menubar.add_cascade(label='File', underline=0, menu=pulldown)


    def editMenu(self):
        pulldown = Menu(self.menubar)
        pulldown.add_command(label='Paste', command=self.notdone)
        pulldown.add_command(label='Spam', command=self.greeting)
        pulldown.add_separator()
        pulldown.add_command(label='Delete', command=self.greeting)
        pulldown.entryconfig(4, state=DISABLED)
        self.menubar.add_cascade(label='Edit', underline=0, menu=pulldown)


    def imageMenu(self):
        photoFiles = ('ora-lp4e.gif', 'pythonPowered.gif',
                      'python_conf_ora.gif')
        pulldown = Menu(self.menubar)
        self.photoObjs = []
        for file in photoFiles:
            img = PhotoImage(file='resources/' + file)
            pulldown.add_command(image=img, command=self.notdone)
            self.photoObjs.append(img)         # созранить ссылку
        self.menubar.add_cascade(label='Image', underline=0, menu=pulldown)


    def greeting(self):
        showinfo('greeting', 'Greetings')


    def notdone(self):
        showerror('Not implemented', 'Not yet available')


    def quit(self):
        if askyesno('Verify quit', 'Are you sure you want to quit?'):
            Frame.quit(self)


if __name__ == '__main__': NewMenuDemo().mainloop() # если запущен как
                                                    # самостоятельный сценарий

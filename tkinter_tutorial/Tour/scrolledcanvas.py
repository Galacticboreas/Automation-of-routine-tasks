"""простой компонент холста в вертикальной прокруткой
"""

from tkinter import *


class ScrolledCanvas(Frame):

    def __init__(self, parent=None, color='brown'):
        Frame.__init__(self, parent)
        self.pack(expand=YES, fill=BOTH)             # сделать растягиваемым
        canv = Canvas(self, bg=color, relief=SUNKEN)
        canv.config(width=300, height=200)           # размер видимой области
        canv.config(scrollregion=(0, 0, 300, 1000))  # углы холста
        canv.config(highlightthickness=0)            # без рамки

        sbar = Scrollbar(self)
        sbar.config(command=canv.yview)              # связать sbar и canv
        canv.config(yscrollcommand=sbar.set)         # сдвиг одного = сдвиг другого
        sbar.pack(side=RIGHT, fill=Y)                # первым добавлен - посл. обрезан
        canv.pack(side=LEFT, expand=YES, fill=BOTH)  # canv обрезается первым

        self.fillContent(canv)
        canv.bind('<Double-1>', self.onDoubleClick)  # установить обр. события
        self.canvas = canv


    def fillContent(self, canv):                     # пререопределить при
        for i in range(10):                          # наследовании
            canv.create_text(150, 50+(i*100), text='spam'+str(i), fill='beige')


    def onDoubleClick(self, event):                  # переопредилить при
        print(event.x, event.y)                      # наследовании
        print(self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))

if __name__ == '__main__': ScrolledCanvas().mainloop()

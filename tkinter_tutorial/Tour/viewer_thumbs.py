"""
отображать все изображения в каталоге в виде кнопок с уменьшенными изображениями,
которые отображают полное изображение при нажатии; требуется PIL для создания
JPEG-файлов и уменьшенных изображений; что нужно сделать: добавьте прокрутку,
если слишком много больших пальцев для окна!
"""

import os, sys, math
from tkinter import *
from PIL import Image                   # <== required for thumbs
from PIL.ImageTk import PhotoImage      # <== required for JPEG display

def makeThumbs(imgdir, size=(100, 100), subdir=''):
    """
    получить уменьшенные изображения для всех изображений в каталоге; для
    каждого изображения создайте и сохраните новый thumb или загрузите и
    верните существующий thumb; при необходимости создает каталог thumb;
    возвращает список (имя файла изображения, объект изображения thumb);
    вызывающий также может запустить listdir в thumb dir для загрузки;
    при неправильных типах файлов может возникнуть ошибка ввода-вывода или
    другое; предостережение: также можно проверить временные метки файлов;
    """
    thumbdir = os.path.join(imgdir, subdir)
    if not os.path.exists(thumbdir):
        os.mkdir(thumbdir)

    thumbs = []
    for imgfile in os.listdir(imgdir):
        thumbpath = os.path.join(thumbdir, imgfile)
        if os.path.exists(thumbpath):
            thumbobj = Image.open(thumbpath)            # use already created
            thumbs.append((imgfile, thumbobj))
        else:
            print('making', thumbpath)
            imgpath = os.path.join(imgdir, imgfile)
            try:
                imgobj = Image.open(imgpath)            # make new thumb
                imgobj.thumbnail(size, Image.ANTIALIAS) # best downsize filter
                imgobj.save(thumbpath)                  # type via ext or passed
                thumbs.append((imgfile, imgobj))
            except:                                     # not always IOError
                print("Skipping: ", imgpath)
    return thumbs

class ViewOne(Toplevel):
    """
    откройте одно изображение во всплывающем окне при создании; фотоизображение
    объект должен быть сохранен: изображения удаляются, если объект восстановлен;
    """
    def __init__(self, imgdir, imgfile):
        Toplevel.__init__(self)
        self.title(imgfile)
        imgpath = os.path.join(imgdir, imgfile)
        imgobj  = PhotoImage(file=imgpath)
        Label(self, image=imgobj).pack()
        print(imgpath, imgobj.width(), imgobj.height())   # размер в пикселях
        self.savephoto = imgobj                           # держи меня в курсе

def viewer(imgdir, kind=Toplevel, cols=None):
    """
    создайте окно ссылок для каталога изображений: по одной кнопке для каждого
    изображения; используйте kind=Tk для отображения в главном окне приложения
    или контейнере фреймов (pack); файл img отличается для каждого цикла: должен
    сохраняться по умолчанию; объекты фотоизображения должны быть сохранены: 
    стирается при восстановлении; рамки для упакованных строк (в отличие от
    сеток, фиксированных размеров, холста); 
    """
    win = kind()
    win.title('Viewer: ' + imgdir)
    quit = Button(win, text='Quit', command=win.quit, bg='beige')   # упаковывай первым,
    quit.pack(fill=X, side=BOTTOM)                                  # чтобы закрепить последним
    thumbs = makeThumbs(imgdir)
    if not cols:
        cols = int(math.ceil(math.sqrt(len(thumbs))))     # fixed or N x N

    savephotos = []
    while thumbs:
        thumbsrow, thumbs = thumbs[:cols], thumbs[cols:]
        row = Frame(win)
        row.pack(fill=BOTH)
        for (imgfile, imgobj) in thumbsrow:
            photo   = PhotoImage(imgobj)
            link    = Button(row, image=photo)
            handler = lambda savefile=imgfile: ViewOne(imgdir, savefile)
            link.config(command=handler)
            link.pack(side=LEFT, expand=YES)
            savephotos.append(photo)
    return win, savephotos

if __name__ == '__main__':
    imgdir = (len(sys.argv) > 1 and sys.argv[1]) or 'resources/'
    main, save = viewer(imgdir, kind=Tk)
    main.mainloop()

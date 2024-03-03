from tkinter import *
from gui7 import HelloPackage   # или from gui7c, где добавлен __getattr__


frm = Frame()
frm.pack()
Label(frm, text='hello').pack()

part = HelloPackage(frm)
part.pack(side=RIGHT)   # НЕУДАЧА! Должно быть part.top.pack(side=RIGHT)
frm.mainloop()

# Traceback (most recent call last):
#   File "tkinter_tutorial\Intro\gui7b.py", line 10, in <module>
#     part.pack(side=RIGHT)   # НЕУДАЧА! Должно быть part.top.pack(side=RIGHT)
# AttributeError: 'HelloPackage' object has no attribute 'pack'

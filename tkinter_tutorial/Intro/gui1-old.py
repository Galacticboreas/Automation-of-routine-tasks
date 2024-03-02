from tkinter import *

Label(None, {'text': 'Hello GUI world!', Pack: {'side': 'top'}}).mainloop()

options = {'text': 'Hello GUI world!'}
layout = {'side': 'top'}
Label(None, **options).pack(**layout)   # ключи должны быть сроками

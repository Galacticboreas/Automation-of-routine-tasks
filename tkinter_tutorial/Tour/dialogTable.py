# определяет таблицу имя:обработчик с демонстрационными примерами

from tkinter.filedialog import askopenfilename # импортировать стандартные
from tkinter.colorchooser import askcolor      # диалоги из Lib\tkinter
from tkinter.messagebox import askquestion, showerror
from tkinter.simpledialog import askfloat


demos = {
    'Open': askopenfilename,
    'Color': askcolor,
    'Query': lambda: askquestion('Warning', 'You typed "rm *"\nConfirm?'),
    'Error': lambda: showerror('Error!', "He's dead, Jim"),
    'Input': lambda: askfloat('Entry', 'Enter credit card number')
}

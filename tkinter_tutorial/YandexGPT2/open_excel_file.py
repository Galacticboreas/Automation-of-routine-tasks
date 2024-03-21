import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd

root = tk.Tk()
root.title("Работа с файлами Excel")

# Функция для открытия файла Excel
def open_file():
    file_path = filedialog.askopenfilename(initialdir="/",
                                           title="Выберите файл", 
                                           filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
    if file_path:
        data = pd.read_excel(file_path)
        display_data(data)

# Функция для отображения данных из файла Excel
def display_data(data):
    frame = ttk.Frame(root)
    frame.grid()

    column_labels = data.columns.values.tolist()
    for col_num, col_label in enumerate(column_labels):
        tk.Label(frame, text=col_label)

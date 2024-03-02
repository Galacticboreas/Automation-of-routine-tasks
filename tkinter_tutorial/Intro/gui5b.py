from gui5 import HelloButton

class MyButton(HelloButton):           # педкласс класса HelloButton
    def callback(self):                # переопределяет метод обработчика
        print('Ignoring press!...')    # события нажатия кнопки

if __name__ == '__main__':
    MyButton(None, text='Hello subclass world').mainloop()

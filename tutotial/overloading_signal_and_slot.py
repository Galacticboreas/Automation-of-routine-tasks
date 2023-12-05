import sys
from PySide6.QtWidgets import QApplication, QPushButton
from PySide6.QtCore import QObject, Signal, Slot


class Communicate(QObject):

    speak = Signal((int,), (str,))

    def __init__(self, parent=None):
        super().__init__(parent)

        self.speak[int].connect(self.say_something)
        self.speak[str].connect(self.say_something)
    
    @Slot(int)
    @Slot(str)
    def say_something(self, arg):
        if isinstance(arg, int):
            print("This is a number:", arg)
        elif isinstance(arg, str):
            print("This is a string:", arg)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    someone = Communicate()

    someone.speak.emit(10)
    someone.speak[str].emit("Hello everybody!")
# This is a number: 10
# This is a string: Hello everybody!

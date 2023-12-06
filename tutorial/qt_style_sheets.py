import sys
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QApplication, QLabel, QWidget, 
                               QListWidgetItem, QListWidget)


class Widget(QWidget):
    def __init__(self, parent=None):
        super(Widget, self).__init__(parent)
    
    menu_widget = QListWidget()
    for i in range(10):
        item = QListWidgetItem(f"Item {i}")
        item.setTextAlignment(Qt.Ali)

if __name__ == "__main__":
    app = QApplication()

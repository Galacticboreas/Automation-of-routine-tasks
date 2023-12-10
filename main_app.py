from operator import imod
import sys

from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QAction, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QLabel,
    QMainWindow,
    QStatusBar,
    QToolBar,
)


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setGeometry(200, 200, 600, 400)
        self.setWindowTitle("Работа с исходными данными")

        self.label = QLabel("Панель инструментов")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.setCentralWidget(self.label)

        self.toolbar = QToolBar("Основная панель инструментов")
        self.toolbar.setIconSize(QSize(16, 16))
        self.addToolBar(self.toolbar)

        self.button_action = QAction(QIcon("resources/icons/android.png"), "&Первая кнопка", self)
        self.button_action.setStatusTip("Это первая кнопка")
        self.button_action.triggered.connect(self.onMyToolBarButtonClick)
        self.button_action.setCheckable(True)
        self.toolbar.addAction(self.button_action)

        self.toolbar.addSeparator()

        self.button_action2 = QAction(QIcon("resources/icons/android.png"), "&Вторая кнопка", self)
        self.button_action2.setStatusTip("Это вторая кнопка")
        self.button_action2.triggered.connect(self.onMyToolBarButtonClick)
        self.button_action2.setCheckable(True)
        self.toolbar.addAction(self.button_action2)

        self.toolbar.addWidget(QLabel("Привет"))
        self.toolbar.addWidget(QCheckBox())

        self.setStatusBar(QStatusBar(self))

        self.menu = self.menuBar()

        self.file_menu = self.menu.addMenu("&Файл")
        self.file_menu.addAction(self.button_action)
        self.file_menu.addSeparator()
        self.file_menu.addAction(self.button_action2)
        
        self.setStatusBar(QStatusBar(self))
    

    def onMyToolBarButtonClick(self, s):
        print("click", s)
    

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())

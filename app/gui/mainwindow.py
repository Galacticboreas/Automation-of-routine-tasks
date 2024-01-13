from PyQt6.QtGui import QAction, QIcon
from PyQt6.QtWidgets import (QMainWindow, QMessageBox, QPushButton,
                             QVBoxLayout, QWidget)

from app.gui.anotherwindow import AnotherWindow


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.window1 = AnotherWindow()
        self.window2 = AnotherWindow()

        layout = QVBoxLayout()
        button1 = QPushButton("Push for Window 1")
        button1.clicked.connect(
            lambda checked: self.toggle_window(self.window1)
        )
        layout.addWidget(button1)

        button2 = QPushButton("Push for Window 2")
        button2.clicked.connect(
            lambda checked: self.toggle_window(self.window2)
        )
        layout.addWidget(button2)

        button_file_open = QAction(
            QIcon("resource/icons/android.png"),
            "&Открыть исходный файл",
            self,
        )
        button_file_open.triggered.connect(self.open_source_file)

        button_collect_db_consumption_per_one_product = QAction(
            "&Собрать БД расход на одно изделие",
            self,
        )
        button_collect_db_consumption_per_one_product.triggered.connect(
            self.collect_db_consumption_per_one_product
        )

        menu = self.menuBar()

        file_menu = menu.addMenu("&File")
        file_menu.addAction(button_file_open)
        file_menu.addAction(button_collect_db_consumption_per_one_product)

        w = QWidget()
        w.setLayout(layout)
        self.setCentralWidget(w)

    def toggle_window(self, window):
        if window.isVisible():
            window.hide()
        else:
            window.show()

    def open_source_file(self, event):
        dlg = QMessageBox(self)
        dlg.setWindowTitle("Вызвать меню открытия файла с исходными данными")
        dlg.setText("Открываем файл с исходными данными")
        self.button_file_open = dlg.exec()

    def collect_db_consumption_per_one_product(self, event):
        dlg = QMessageBox(self)
        dlg.setWindowTitle("Собрать БД материалы расх. на 1 изд.")
        dlg.setText("Открываем файл с выгрузкой данных")
        self.button_collect_db_consumption_per_one_product = dlg.exec()

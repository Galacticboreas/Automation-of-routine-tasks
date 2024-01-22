from PyQt6 import QtCore, QtWidgets
from PyQt6.QtWidgets import QMainWindow, QWidget, QPushButton
from PyQt6.QtCore import QSize
from PyQt6.QtGui import QIcon, QAction

class MainWindow(QMainWindow):
    
    def __init__(self):
        QMainWindow.__init__(self)

        self.setMinimumSize(QSize(300, 100))
        self.setWindowTitle("Автоматизация рутинных задач")

        # Create new action
        newAction = QAction(QIcon('new.png'), '&New', self)        
        newAction.setShortcut('Ctrl+N')
        newAction.setStatusTip('New document')
        newAction.triggered.connect(self.newCall)

        # Create new action
        openAction = QAction(QIcon('open.png'), '&Open', self)        
        openAction.setShortcut('Ctrl+O')
        openAction.setStatusTip('Open document')
        openAction.triggered.connect(self.openCall)

        # Create exit action
        exitAction = QAction(QIcon('exit.png'), '&Exit', self)        
        exitAction.setShortcut('Ctrl+Q')
        exitAction.setStatusTip('Exit application')
        exitAction.triggered.connect(self.exitCall)

        # Create menu bar and add action
        menuBar = self.menuBar()
        fileMenu = menuBar.addMenu('&File')
        fileMenu.addAction(newAction)
        fileMenu.addAction(openAction)
        fileMenu.addAction(exitAction)

    def openCall(self):
        print('Open')

    def newCall(self):
        print('New')

    def exitCall(self):
        print('Exit app')

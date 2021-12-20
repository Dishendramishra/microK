from PyQt5 import QtWidgets, uic
from qt_material import *

class RuntimeStylesheets(QtWidgets.QMainWindow, QtStyleTools):

    def __init__(self):
        super().__init__()
        self.main = QUiLoader().load('main_window.ui', self)

        self.show_dock_theme(self.main)


var = RuntimeStylesheets()
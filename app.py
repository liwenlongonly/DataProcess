# -*- coding: utf-8 -*-
from PyQt5.QtWidgets import QApplication
from PyQt5 import QtGui
import sys
from main_window import MainWindow

class MyApplication(QApplication):

    def __init__(self, argv):
        super().__init__(argv)
        self.setWindowIcon(QtGui.QIcon("home.jpeg"))


if __name__ == '__main__':
    app = MyApplication(sys.argv)
    window = MainWindow()
    window.init_window()
    window.show()
    sys.exit(app.exec_())
# -*- coding: utf-8 -*-
from PyQt5.QtWidgets import QApplication
import sys
from main_window import MainWindow


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.init_window()
    window.show()
    sys.exit(app.exec_())
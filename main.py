import sys
from PyQt5 import QtWidgets
from ui_mainwindow import MainWindow
import sympy
import plotly
import os
import pymysql
from ERRORS import NETWORK_ERR

if __name__ == '__main__':
    # print(type(os.getcwd()))
    app = QtWidgets.QApplication(sys.argv)
    try:
        myWin = MainWindow()
        myWin.show()
        sys.exit(app.exec_())
    except NETWORK_ERR:
        app.quit()

# pyuic5 -o mainwindow.py mainwindow.ui
# pyinstaller -F main.py --noconsole
# pyinstaller -D main.py --noconsole

# 加载spec里面
# import sys
# sys.setrecursionlimit(5000)
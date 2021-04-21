from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.Qt import *
# from database import Database
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QCheckBox, QGridLayout, QTreeWidgetItem, QLabel, QLineEdit, QDialog, QMessageBox
from Ui_calculating import Ui_Dialog
import os
import win32api

class calculating(QtWidgets.QDialog, Ui_Dialog):
    def __init__(self, parent = None):
        super().__init__(parent)
        self.setupUi(self)
        QtWidgets.QApplication.setStyle('Fusion')  # ui风格为Fusion
        self.setWindowTitle("正在计算，请稍候")
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)  # 无标题栏
        self.percentage = 0
        self.progressBar.setValue(0)
        self.txt_lb.setText("已完成0%")

    def setpercentage(self, i):
        self.percentage = int(i * 100)
        self.txt_lb.setText("已完成%d" % self.percentage + "%")
        self.progressBar.setValue(self.percentage)

    def finish_cal(self, citypath):
        self.city = citypath
        self.resize(400, 250)
        self.label.setText("计算已完成")
        self.txt_lb.deleteLater()
        self.progressBar.deleteLater()
        self.lb_finish = QLabel(self)
        self.lb_finish.setText("计算已完成！感谢你的耐心等待。")
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(10)
        self.lb_finish.setFont(font)
        self.lb_finish.setGeometry(50, 80, 300, 20)
        self.display_path = QLabel(self)
        self.display_path.setGeometry(50, 130, 300, 60)
        self.display_path.setWordWrap(True)
        self.display_path.setAlignment(QtCore.Qt.AlignTop)
        self.display_path.setText("请在\"%s\"下查看输出文件。" % citypath)
        self.display_path.setFont(font)
        self.btnOk = QPushButton(self)
        self.btnOk.setText("确定")
        self.btnOk.setFont(font)
        self.btnOk.setGeometry(85, 200, 93, 28)
        self.btnOk.clicked.connect(self.close_ui)
        self.btnopen = QPushButton(self)
        self.btnopen.setGeometry(210, 200, 130, 28)
        self.btnopen.setText("在文件夹中显示")
        self.btnopen.setFont(font)
        self.btnopen.clicked.connect(self.show_in_folder)
        self.lb_finish.show()
        self.btnOk.show()
        self.display_path.show()
        self.btnopen.show()
    
    def close_ui(self):
        self.close()

    def show_in_folder(self):
        win32api.ShellExecute(0, 'open', self.city, '', '', 1)
        # print(self.city)
        self.close()

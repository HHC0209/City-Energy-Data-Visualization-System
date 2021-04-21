from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtGui import QMovie
# from mainwin import mainwin

class loading_win(QWidget):
    def __init__(self, mainWin):
        super(loading_win, self).__init__()
        # 获取主窗口的坐标
        self.m_winX = mainWin.x()
        self.m_winY = mainWin.y()
        self.w = mainWin.width()
        self.h = mainWin.height()
        self.percentage = 0
        self.initUI()
        self.setWindowModality(Qt.ApplicationModal)
        # mainWin.setWindowModality(Qt.ApplicationModal)
        

    def setpercentage(self, i):
        # print(i)
        self.percentage = i
        x = self.percentage * 100
        self.txtlb.setText("正在计算，已完成%d" % x + "%")

    def initUI(self):
        # 设置窗口基础类型
        # self.resize(self.w, self.h)	# 设置加载界面的大小
        # self.move(self.m_winX+340,self.m_winY+155)	# 移动加载界面到主窗口的中心
        self.setGeometry(self.m_winX, self.m_winY + 30, self.w, self.h - 30)
        self.setWindowFlags(Qt.SubWindow | Qt.FramelessWindowHint)
        # self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog | Qt.WindowStaysOnTopHint) # 设置窗口无边框|对话框|置顶模式
        # 设置背景透明
        # self.setAttribute(Qt.WA_TranslucentBackground)
        self.setWindowOpacity(0.8)
        # 加载动画画面
        self.loading_gif = QMovie('loading.gif')
        self.loading_label = QLabel(self)
        self.loading_label.setGeometry(218, 150, 64, 64)
        self.loading_label.setMovie(self.loading_gif)
        # self.loading_label.setMaximumSize(64, 64)
        # self.loading_label.setMinimumSize(64, 64)
        self.loading_gif.start()
        self.txtlb = QLabel(self)
        self.txtlb.setGeometry(50, 300, 400, 50)
        self.txtlb.setText("正在计算，已完成%d" % self.percentage + "%")
        self.txtlb.setAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignVCenter)
        self.txtlb.show()
        self.btnKill = QPushButton(self)
        self.btnKill.setGeometry(200, 400, 100, 50)
        self.btnKill.setText("停止")
        self.btnKill.show()
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.Qt import *
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QCheckBox, QGridLayout, QTreeWidgetItem, QLabel, QLineEdit, QDialog, QMessageBox
from Ui_detail_dialog import Ui_Dialog
from ui_dialog_1 import Dialog_1
from ERRORS import NETWORK_ERR, INSERT_FAILURE, UPDATE_FAILURE, EXECUTE_FAILURE

# 查询显示数据详细信息的窗口逻辑
class detail_dialog(QtWidgets.QDialog, Ui_Dialog):
    def __init__(self, id, item, parent = None):
        super().__init__(parent)
        self.setupUi(self)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)  # 无标题栏
        self.id = id
        self.item = item
        self.title = ''
        self.info = ''
        self.detail_path = []

        # 设置控件逻辑
        self.btnOk.clicked.connect(self.ok_clicked)
        self.btn_close.clicked.connect(self.close_clicked)
        self.tbw_info.setEditTriggers(QtWidgets.QTableView.NoEditTriggers)


    def setTitle(self, title):
        self.title = title + "查询详情"
        self.titleLb.setText(self.title)
        self.titleLb.setWordWrap(True)


    def set_detail_path(self):
        """
        设置当前选中的数据项的详细路径（由级别1至6等组成）
        """
        name = self.item.name[-1]
        for item in self.item.data:
            cnt = 0
            for i in range(1, 7):
                query = "级别%d" % i
                if item[query] == name:
                    cnt = i
                    break
            if cnt == 6:
                self.detail_path.append(self.item.realname)
            else:
                cnt1 = cnt + 1
                query = "级别%d" % cnt1
                tmp = item[query]
                if cnt1 + 1 <= 6:
                    for i in range(cnt1 + 1, 7):
                        query = "级别%d" % i
                        j = i - 1
                        post_query = "级别%d" % j
                        if item[query] != item[post_query]:
                            tmp += ", " + item[query]
                self.detail_path.append(tmp)


    def table_display(self): 
        """
        初始化表格内容，显示表格
        """
        self.tbw_info.clear()
        self.tbw_info.setRowCount(len(self.item.data))
        self.header = ['明细']
        for i in self.item.year:
            self.header.append(i)

        self.tbw_info.setColumnCount(len(self.header))
        self.tbw_info.setHorizontalHeaderLabels(self.header)
        self.tbw_info.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)

        self.set_detail_path()
        for i in range(0, len(self.detail_path)):
            item = QtWidgets.QTableWidgetItem(self.detail_path[i])
            item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            self.tbw_info.setItem(i, 0, item)

        for i in range(1, len(self.header)):
            year = self.tbw_info.horizontalHeaderItem(i).text()
            for j in range(0, len(self.item.data)):
                try:
                    res = self.item.data[j][year]
                except:
                    res = ""
                    
                item = QtWidgets.QTableWidgetItem(res)
                item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                self.tbw_info.setItem(j, i, item)



    def ok_clicked(self):
        self.close()
    
    def close_clicked(self):
        self.close()

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.m_flag = True
            self.m_Position = event.globalPos() - self.pos()
            event.accept()
            self.setCursor(QtGui.QCursor(QtCore.Qt.OpenHandCursor))

    def mouseMoveEvent(self, QMouseEvent):
        if QtCore.Qt.LeftButton and self.m_flag:
            self.move(QMouseEvent.globalPos() - self.m_Position)
            QMouseEvent.accept()

    def mouseReleaseEvent(self, QMouseEvent):
        self.m_flag = False
        self.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))

from Ui_search_add import Ui_Dialog
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import QMessageBox

# 搜索添加页面的相关逻辑
class search_add(QtWidgets.QDialog, Ui_Dialog):
    def __init__(self, header, data, parent = None):
        super().__init__(parent)
        self.setupUi(self)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint) 
        self.header = header
        self.data = data
        self.setUp_table()
        self.lab_example.setText('')

        # 几个按钮的逻辑
        self.btnCityOk.clicked.connect(self.onBtnCityOk)
        self.btnOk.clicked.connect(self.onBtnOk)
        self.btnCancel.clicked.connect(self.onBtnCancel)


    def setUp_table(self):
        """
        初始化表格，显示窗口中tablewidget的内容
        """

        self.tbw_dialog.setRowCount(1)
        self.tbw_dialog.setColumnCount(len(self.header))
        self.tbw_dialog.setHorizontalHeaderLabels(self.header)
        self.tbw_dialog.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)

        index = self.data[-1]
        item = QtWidgets.QTableWidgetItem(str(index))
        item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        self.tbw_dialog.setItem(0, 0, item)

        item = QtWidgets.QTableWidgetItem('')
        item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        self.tbw_dialog.setItem(0, 1, item)

        for i in range(len(self.data) - 1):
            data_item = self.data[i]
            item = QtWidgets.QTableWidgetItem(data_item)
            item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            self.tbw_dialog.setItem(0, i + 2, item)

        for i in range(8, len(self.header)):
            item = QtWidgets.QTableWidgetItem('0.0')
            item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            self.tbw_dialog.setItem(0, i, item)

        # 设置第一个单元格（编码）不可编辑
        self.tbw_dialog.setItemDelegateForColumn(0, EmptyDelegate(self))


    def onBtnCityOk(self):
        """
        确定选择城市按钮的逻辑
        """
        text = self.le_city.text()
        if text == '':
            self.lab_example.setText("<font color='red'>   请输入城市。</font>")
        else:
            item = QtWidgets.QTableWidgetItem(text)
            self.tbw_dialog.setItem(0, 1, item)

    def onBtnOk(self):
        """
        确定按钮的逻辑
        """
        empty = []
        for index in range(0, 8):
            if not self.tbw_dialog.item(0, index).text():
                empty.append(self.header[index])

        if empty:
            temp = ''
            for item in empty:
                temp += (item + ' ')
            self.lab_example.setText("<font color='red'>   %s不能为空！</font>" % temp)
        else:
            self.accept()


    def onBtnCancel(self):
        self.reject()


    def get_data(self):
        """
        获取表格中的所有数据
        """
        data = []
        for i in range(len(self.header)):
            item = self.tbw_dialog.item(0, i)
            data.append(item.text())
        return data

class EmptyDelegate(QItemDelegate):
    def __init__(self,parent):
        super(EmptyDelegate, self).__init__(parent)
 
    def createEditor(self, QWidget, QStyleOptionViewItem, QModelIndex):
        return None
from PyQt5 import QtCore, QtGui, QtWidgets
from dialog_1 import Ui_Dialog_1


class Dialog_1(QtWidgets.QDialog, Ui_Dialog_1):
    def __init__(self, mode, header, data=[], parent=None):
        super(Dialog_1, self).__init__(parent)
        self.setupUi(self)

        self.data = data
        self.header = header
        self.mode = mode
        if self.mode == 0:
            self.hint = "增添数据"
        elif self.mode == 1:
            self.hint = "修正数据——与此路径相关的所有数据已全部列出如下表\n请在表格原位双击进行修改"

        self.lab_example.setText(self.hint)  # 设置lab_example显示提示
        QtWidgets.QApplication.setStyle('Fusion')  # ui风格为Fusion
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)  # 无标题栏
        self.initiate_tbw()

# ======================================================自定义方法======================================================
    def initiate_tbw(self):
        # self.tbw_dialog.setRowCount(1)
        if self.mode == 1:
            self.tbw_dialog.setRowCount(len(self.data))
        else:
            self.tbw_dialog.setRowCount(1)

        self.tbw_dialog.setColumnCount(len(self.header))
        self.tbw_dialog.setHorizontalHeaderLabels(self.header)
        self.tbw_dialog.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)

        if self.data:
            # print(self.data)
            for i in range(len(self.data)):
                for index in range(len(self.header)):
                    item = QtWidgets.QTableWidgetItem(str(self.data[i][index]))
                    item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                    # if index in [0, 1, 2, 3, 4, 5, 6, 7]:
                    #     item.setFlags(QtCore.Qt.ItemIsEditable)
                    self.tbw_dialog.setItem(i, index, item)

        else:
            for index in range(len(self.header)):
                if index in [0, 1, 2, 3, 4, 5, 6, 7]:
                    item = QtWidgets.QTableWidgetItem('')
                    item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                    self.tbw_dialog.setItem(0, index, item)
                else:
                    item = QtWidgets.QTableWidgetItem('0.0')
                    item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                    self.tbw_dialog.setItem(0, index, item)



    def get_data1(self):
        data = []
        for index in range(len(self.header)):
            data.append(self.tbw_dialog.item(0, index).text())
        return data

    def get_data(self):
        data = []
        for i in range(len(self.data)):
            tmp = []
            for j in range(len(self.header)):
                tmp.append(self.tbw_dialog.item(i, j).text())
            data.append(tmp)
        return data

# ======================================================自定义方法======================================================

# ======================================================UI组件方法======================================================
    @QtCore.pyqtSlot()
    def on_btn_ok_clicked(self):
        empty = []
        for index in range(0, 8):
            if not self.tbw_dialog.item(0, index).text():
                empty.append(self.header[index])

        if empty:
            temp = ''
            for item in empty:
                temp += (item + ' ')
            self.lab_example.setText(self.hint + "<font color='red'>   %s不能为空！</font>" % temp)
        else:
            self.accept()

    @QtCore.pyqtSlot()
    def on_btn_cancel_clicked(self):
        self.reject()

    # 拖动窗口
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
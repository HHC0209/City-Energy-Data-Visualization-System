# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'f:\pyqt5\project_developing\dialog_1.ui'
#
# Created by: PyQt5 UI code generator 5.15.2
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Dialog_1(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(1200, 400)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(10)
        Dialog.setFont(font)
        Dialog.setStyleSheet("background-color: rgb(50, 50, 50);\n"
"color: rgb(255, 255, 255);")
        self.btn_cancel = QtWidgets.QPushButton(Dialog)
        self.btn_cancel.setGeometry(QtCore.QRect(1110, 360, 80, 25))
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(10)
        self.btn_cancel.setFont(font)
        self.btn_cancel.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btn_cancel.setStyleSheet("color: rgb(255, 255, 255);\n"
"background-color: rgb(64, 65, 66);")
        self.btn_cancel.setObjectName("btn_cancel")
        self.btn_ok = QtWidgets.QPushButton(Dialog)
        self.btn_ok.setGeometry(QtCore.QRect(1020, 360, 80, 25))
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(10)
        self.btn_ok.setFont(font)
        self.btn_ok.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.btn_ok.setStyleSheet("color: rgb(255, 255, 255);\n"
"background-color: rgb(64, 65, 66);")
        self.btn_ok.setObjectName("btn_ok")
        self.lab_example = QtWidgets.QLabel(Dialog)
        self.lab_example.setGeometry(QtCore.QRect(10, 5, 880, 60))
        self.lab_example.setMinimumSize(QtCore.QSize(880, 60))
        self.lab_example.setMaximumSize(QtCore.QSize(880, 60))
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.lab_example.setFont(font)
        self.lab_example.setText("")
        self.lab_example.setObjectName("lab_example")
        self.tbw_dialog = QtWidgets.QTableWidget(Dialog)
        self.tbw_dialog.setGeometry(QtCore.QRect(10, 70, 1180, 270))
        self.tbw_dialog.setStyleSheet("background-color: rgb(54, 54, 54);\n"
"color: rgb(255, 255, 255);")
        self.tbw_dialog.setObjectName("tbw_dialog")
        self.tbw_dialog.setColumnCount(0)
        self.tbw_dialog.setRowCount(0)

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.btn_cancel.setText(_translate("Dialog", "取消"))
        self.btn_ok.setText(_translate("Dialog", "确定"))

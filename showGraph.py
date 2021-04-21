from PyQt5 import QtCore, QtGui, QtWidgets
from Ui_showGraph import Ui_Dialog

class showGraph(QtWidgets.QDialog, Ui_Dialog):
    def __init__(self, lab_draw, parent = None):
        super().__init__(parent)

        self.setupUi(self)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint) 

        self.widget = lab_draw
        self.bnt_close.clicked.connect(self.closeWin)


    def display(self):
        self.sca_plot.setWidget(self.widget)


    def closeWin(self):
        self.close()
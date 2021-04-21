from PyQt5 import QtCore, QtGui, QtWidgets
from dialog_2 import Ui_Dialog_2
import xlrd


class Dialog_2(QtWidgets.QDialog, Ui_Dialog_2):
    def __init__(self, parent=None):
        super(Dialog_2, self).__init__(parent)
        self.setupUi(self)

        QtWidgets.QApplication.setStyle('Fusion')  # ui风格为Fusion
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)  # 无标题栏
        self.tbw_draw.setEditTriggers(QtWidgets.QTableView.NoEditTriggers)  # 禁止编辑
        self.tbw_draw.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)

# ======================================================自定义方法======================================================
    def display(self, draw_excel_path):
        workbook = xlrd.open_workbook(draw_excel_path)
        worksheet = workbook.sheet_by_index(0)
        header = worksheet.row_values(0)  # 读取excel第1行的header
        if header[0] == '编码':
            #  读取excel中有多少个不同的region
            regions = worksheet.col_values(1)
            exist_regions = []
            for region in regions[1:]:
                if region not in exist_regions and region:
                    exist_regions.append(region)
            region_str = str(exist_regions).replace("'", "").replace("[", "").replace("]", "")
            self.display_0(worksheet)
            self.description = '%s数据' % region_str
        else:
            self.display_1(worksheet)
            # self.description = '%d%s 能流图数据' % (year, region)
            self.description = "能流图"

        self.lab_description.setText(self.description)

    def display_0(self, worksheet):
        header = worksheet.row_values(0)  # 读取excel第1行的header

        self.tbw_draw.setRowCount(worksheet.nrows - 1)
        self.tbw_draw.setColumnCount(worksheet.ncols)
        self.tbw_draw.setHorizontalHeaderLabels(header)

        # 数据从excel的第2行开始
        for row_num in range(1, worksheet.nrows):
            row = worksheet.row_values(row_num)
            for col_num in range(len(row)):
                if col_num == 0:
                    item = QtWidgets.QTableWidgetItem(str(int(row[col_num])))
                elif col_num >= 8:
                    try:
                        item = QtWidgets.QTableWidgetItem(str(round(float(row[col_num]), 3)))
                    except:
                        item = QtWidgets.QTableWidgetItem(str(0))
                else:
                    item = QtWidgets.QTableWidgetItem(str(row[col_num]))
                self.tbw_draw.setItem(row_num - 1, col_num, item)

    def display_1(self, worksheet):
        self.tbw_draw.setRowCount(worksheet.nrows - 1)
        self.tbw_draw.setColumnCount(worksheet.ncols)
        
        header = worksheet.row_values(0)
        self.tbw_draw.setHorizontalHeaderLabels(header)
        for row_num in range(1, worksheet.nrows):
            row = worksheet.row_values(row_num)
            for i in range(len(row)):
                item_val = row[i]
                # print(item_val)
                item = QtWidgets.QTableWidgetItem(str(item_val))
                self.tbw_draw.setItem(row_num - 1, i, item)

    def display_2(self, data):
        self.tbw_draw.setRowCount(len(data))
        if not len(data):
            return 
        col = len(data[0].keys())
        self.tbw_draw.setColumnCount(col)
        header = [item for item in data[0].keys()]
        self.tbw_draw.setHorizontalHeaderLabels(header)
        for rn in range(len(data)):
            for cn in range(col):
                val = data[rn][header[cn]]
                item = QtWidgets.QTableWidgetItem(str(val))
                self.tbw_draw.setItem(rn, cn, item)


# ======================================================自定义方法======================================================

# ======================================================UI组件方法======================================================
    @QtCore.pyqtSlot()
    def on_btn_ok_clicked(self):
        self.accept()

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
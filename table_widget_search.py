from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.Qt import *
from tbw_item import tbw_item
from detail_dialog import detail_dialog
from ui_dialog_1 import Dialog_1
import xlwt


class TableWidgetSearch:
    def __init__(self, widget):
        self.datas = []    #包含了表格的各行，类型为tbw_item。
        self.selected_year = []
        self.widget = widget
        self.contain = []
        self.modify_for_row = []
        self.detail_for_row = []


    def view_detail(self, id):
        dialog_detail = detail_dialog(id = id, item = self.datas[id])
        name = self.datas[id].realname
        dialog_detail.setTitle(name)
        dialog_detail.setInfo("Here are some information.")
        dialog_detail.exec_()


    def buttonForRow(self, id):
        """
        设置表格中每一行的两个按钮
        :param id: 行号
        """
        widget = QWidget()
        detailBtn = QPushButton("查看详情")
        detailBtn.setStyleSheet("color: rgb(255, 255, 255); background-color: rgb(75, 225, 75); font-family:'微软雅黑'; font-Point Size:10")
        detailBtn.clicked.connect(lambda:self.view_detail(id))
        self.detail_for_row.append(detailBtn)

        modifyBtn = QPushButton("修正")
        modifyBtn.setStyleSheet("color: rgb(255, 255, 255); background-color: rgb(225, 174, 0); font-family:'微软雅黑'; font-Point Size:10")
        self.modify_for_row.append(modifyBtn)
        modifyBtn.clicked.connect(lambda:self.modify(id))
        hLayout = QHBoxLayout()
        hLayout.addWidget(detailBtn)
        hLayout.addWidget(modifyBtn)
        hLayout.setContentsMargins(5,2,5,2)
        widget.setLayout(hLayout)

        return widget


    def clear(self):
        self.widget.clear()
        self.datas = []
        self.selected_year = []
        self.widget.setRowCount(0)
        self.widget.setColumnCount(0)


    def add_data(self, datas, selected_years):
        self.datas = datas
        self.selected_year = selected_years
        for i in range(len(self.datas)):
            datas[i].setRow(i)
            datas[i].setRealName()


    def add_data_dialog(self, data):
        flag = -1
        for index, d in enumerate(self.datas):
            if d['编码'] == data[0] and d['地域'] == data[1]:
                flag = index

        if flag >= 0:
            # 6个级别 从第2列开始
            for cat_index in range(2, 8):
                self.datas[flag]['级别%d' % (cat_index - 1)] = data[cat_index]
            # 从第8列开始
            for year_index in range(len(self.selected_year)):
                value = data[8+year_index]
                self.datas[flag][self.selected_year[year_index]] = value
        else:
            temp = {}
            for i in range(len(self.header)):
                temp[self.header[i]] = data[i]
            self.datas.append(temp)


    def delete_data(self, index, region):
        for i in range(len(self.datas)):
            if self.datas[i]['编码'] == index and self.datas[i]['地域'] == region:
                del self.datas[i]
                break


    def display(self):
        """
        显示表格内容
        """
        self.widget.clear()
        self.widget.setRowCount(len(self.datas))
        self.widget.setColumnCount(2)
        self.header = ['路径','操作']
        self.widget.setHorizontalHeaderLabels(self.header)
        self.widget.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)

        row = 0
        col = 0
        for point in self.datas:
            item = QtWidgets.QTableWidgetItem(point.realname)
            item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            self.widget.setItem(row, col, item)
            row += 1


    def export_excel(self, filepath):
        workbook = xlwt.Workbook(encoding='utf-8')
        worksheet = workbook.add_sheet('Data')

        style_hv = xlwt.XFStyle()
        font = xlwt.Font()
        font.name = '微软雅黑'

        alignment_hv = xlwt.Alignment()
        alignment_hv.horz = xlwt.Alignment.HORZ_CENTER
        alignment_hv.vert = xlwt.Alignment.VERT_CENTER

        style_hv.alignment = alignment_hv
        style_hv.font = font

        worksheet.write(0, 0, '编码', style_hv)
        worksheet.write(0, 1, '地域', style_hv)
        for i in range(1, 7):
            worksheet.write(0, 1+i, '级别%d' % i, style_hv)
        for year_index in range(len(self.selected_year)):
            worksheet.write(0, 8+year_index, self.selected_year[year_index], style_hv)

        # sum = 0
        # for item in self.datas:
        #     sum += len(item.data)
        # print(sum)

        row = 1
        for item in self.datas:
            for record in item.data:
                col = 0
                for key in record.keys():
                    if key == 'time':
                        continue
                    if key in ['2010', '2011', '2012', '2013', '2014', '2015', '2016', '2017', '2018', '2019']:
                        # print(key)
                        if key in self.selected_year:
                            data = record[key]
                        else:
                            continue
                    else:
                        data = record[key]
                    try:
                        data = float(data)
                    except ValueError:
                        pass
                    worksheet.write(row, col, data, style_hv)
                    col += 1
                row += 1
        # for row in range(sum):
        # # for row in range(len(self.datas)):
        #     for col in range(len(self.datas[0].data[0].keys())-1):
        #         data = self.widget.item(row, col).text()
        #         try:
        #             data = float(data)
        #         except ValueError:
        #             pass
        #         worksheet.write(row+1, col, data, style_hv)

        worksheet.col(0).width = 3000
        worksheet.col(1).width = 3000
        for col in range(2, 8):
            worksheet.col(col).width = 6200
        for col in range(8, 8 + len(self.selected_year) + 1):
        # for col in range(8, len(self.datas[0].keys())-1):
            worksheet.col(col).width = 3000
        workbook.save(filepath)

    def refresh(self, id):
        """
        刷新表格内容
        """
        item = QtWidgets.QTableWidgetItem(self.datas[id].realname)
        item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
        self.widget.setItem(id, 0, item)

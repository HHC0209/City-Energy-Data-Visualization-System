from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QCheckBox, QGridLayout, QTreeWidgetItem, QLabel, QLineEdit, QDialog, QMessageBox
from PyQt5.Qt import Qt
from PyQt5.QtCore import QThread
from PyQt5.Qt import *
from mainwindow import Ui_MainWindow
from ui_dialog_1 import Dialog_1
from ui_dialog_2 import Dialog_2
from draw import Draw
from database1 import Database
from table_widget_search import TableWidgetSearch
from table_widget_process import TableWidgetProcess
import time
import xlrd
import win32api
import PyQt5.sip as sip
import os
import json
import shutil
import importlib
import numpy as np
from tbw_item import tbw_item
from detail_dialog import detail_dialog
from dialog_search_result import dialog_search_result
from search_modify import search_modify
from search_add import search_add
from loading_win import loading_win
from auto import auto
from calculating import calculating
# import NETWORK_ERR
from ERRORS import NETWORK_ERR, INSERT_FAILURE, UPDATE_FAILURE, EXECUTE_FAILURE

class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    signal = pyqtSignal(float)
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)

        try:
            self.database = Database()  # 初始化数据库
        except NETWORK_ERR as net_err:
            QMessageBox.critical(self, "警告", net_err.__str__() + "请检查网络后重启应用。")
            raise NETWORK_ERR
            return

        self.year = []  # 存放当前数据库中拥有的年份
        self.region = []  # 存放当前数据库中拥有的地域
        self.category = {}  # 存放当前数据库中拥有的指标
        self.category_data = []  # 存放当前数据库中拥有的指标数据
        self.forecast_funcs = []  # 存放当前lsw_forecast中拥有的预测函数文件的路径
        self.analyze_funcs = []  # 存放当前lsw_analyze中拥有的分析函数文件的路径
        self.plot_funcs = []  # 存放当前拥有的绘图函数文件的路径
        self.templates = [] # 存放当前的模板
        self.contain = []
        self.contain_tbl = []
        self.search_path = []
        self.draw_data = []

        QtWidgets.QApplication.setStyle('Fusion')  # ui风格为Fusion
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)  # 无标题栏
        self.lab_draw = QLabel()
        self.sca_plot_draw_2.setWidget(self.lab_draw)
        self.draw_excel_path = ''  # 存放绘图中导入excel的文件位置
        self.draw_class = Draw(self.lab_draw)
        self.draw_class.signal.connect(self.use_database)
        self.usedb = True
        self.initialize_file()  # 初始化软件的文件

        self.category_search_checked = False
        self.category_process_checked = False
        
        self.category_draw_checked = False
     
        self.glo_city_search = QGridLayout(self.sawc_city_search)
        self.glo_year_search = QGridLayout(self.sawc_year_search)
        
        self.glo_city_process = QGridLayout(self.sawc_city_process)
        
        self.glo_year_process = QGridLayout(self.sawc_year_process)
        
        self.glo_city_draw = QGridLayout(self.sawc_city_process_2)
       
        self.glo_year_draw = QGridLayout(self.sawc_year_process_2)


        self.initialize_template_list()
        self.refresh()

        self.tbw_search.setEditTriggers(QtWidgets.QTableView.NoEditTriggers)  # 禁止编辑
        self.tbw_search.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.tbw_search.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)  # 只有行能被选中
        self.tbw_search_class = TableWidgetSearch(self.tbw_search)

        self.tbw_process.setEditTriggers(QtWidgets.QTableView.NoEditTriggers)  # 禁止编辑
        self.tbw_process_class = TableWidgetProcess(self.tbw_process)

        self.lsw_forecast.setEditTriggers(QtWidgets.QTableView.NoEditTriggers)  # 禁止编辑
        self.lsw_forecast.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.initialize_listWidget(self.lsw_forecast, self.forecast_funcs)  # 初始化listWidget
        self.lsw_analyze.setEditTriggers(QtWidgets.QTableView.NoEditTriggers)  # 禁止编辑
        self.lsw_analyze.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.initialize_listWidget(self.lsw_analyze, self.analyze_funcs)  # 初始化listWidget
        self.lsw_plot_2.setEditTriggers(QtWidgets.QTableView.NoEditTriggers)  # 禁止编辑
        self.lsw_plot_2.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.initialize_listWidget(self.lsw_plot_2, self.plot_funcs)  # 初始化listWidget
        self.btn_search.clicked.connect(self.search)
        self.le_search.setPlaceholderText("请输入关键字搜索")
        self.btnbrowse1.clicked.connect(self.browse_path)
        self.btnbrowse2.clicked.connect(self.browse_output)
        self.lineEdit_threshold.setText('0.535')
        self.btn_refresh.clicked.connect(self.refresh)

# ======================================================自定义方法======================================================

    def use_database(self):
        self.usedb = False

    def refresh(self):
        """
        刷新主界面
        """
        self.form_year_from_db()
        self.form_region_from_db()
        self.form_category_structure_from_db()
        self.clear_gridLayout(self.glo_city_search)  # 清空gridLayout
        self.clear_gridLayout(self.glo_year_search)  # 清空gridLayout
        self.clear_gridLayout(self.glo_city_process)  # 清空gridLayout
        self.clear_gridLayout(self.glo_year_process)  # 清空gridLayout
        self.clear_gridLayout(self.glo_city_draw)  # 清空gridLayout
        self.clear_gridLayout(self.glo_year_draw)  # 清空gridLayout
        self.city_search_checked = False
        self.initialize_gridLayout(self.glo_city_search, self.region)  # 初始化gridLayout
        self.year_search_checked = False
        self.initialize_gridLayout(self.glo_year_search, self.year)  # 初始化gridLayout
        self.city_process_checked = False
        self.initialize_gridLayout(self.glo_city_process, self.region)  # 初始化gridLayout
        self.year_process_checked = False
        self.initialize_gridLayout(self.glo_year_process, self.year)  # 初始化gridLayout

        self.city_draw_checked = False
        # self.glo_city_draw = QGridLayout(self.sawc_city_process_2)
        self.initialize_gridLayout(self.glo_city_draw, self.region)  # 初始化gridLayout
        self.year_draw_checked = False
        # self.glo_year_draw = QGridLayout(self.sawc_year_process_2)
        self.initialize_gridLayout(self.glo_year_draw, self.year)  # 初始化gridLayout
        self.trw_search.clear()
        self.trw_process.clear()
        self.trw_process_2.clear()
        self.initiate_treeWidget(self.trw_search)
        self.initiate_treeWidget(self.trw_process)
        self.initiate_treeWidget(self.trw_process_2)


    # 第一次运行时初始化软件的文件 以后则读取文件
    def initialize_file(self):
        dirs = ["funcs", "funcs/forecast_funcs", "funcs/analyze_funcs", "funcs/plot_funcs", "graphs", "configs", "templates"]
        for d in dirs:
            if not os.path.exists(d):
                os.mkdir(d)

        # 第一次运行时创建year.cfg 用来储存self.year中新添加的项 以后则读取文件 储存在self.year中
        if not os.path.exists("configs/year.cfg"):
            with open("configs/year.cfg", mode="w", encoding="utf-8") as f:
                pass
        else:
            with open("configs/year.cfg", mode="r", encoding="utf-8") as f:
                year_lines = f.readlines()
                self.year = [year_line.strip() for year_line in year_lines]
                # 对年份进行排序 （whh）
                self.year.sort()

        # 第一次运行时创建region.cfg 用来储存self.region中新添加的项 以后则读取文件 储存在self.region中
        if not os.path.exists("configs/region.cfg"):
            with open("configs/region.cfg", mode="w", encoding="utf-8") as f:
                pass
        else:
            with open("configs/region.cfg", mode="r", encoding="utf-8") as f:
                region_lines = f.readlines()
                self.region = [region_line.strip() for region_line in region_lines]
        # 第一次运行时创建category.cfg 用来储存self.category中新添加的项 以后则读取文件 储存在self.category中
        if not os.path.exists("configs/category.json"):
            with open("configs/category.json", mode="w", encoding="utf-8") as f:
                pass
        else:
            with open("configs/category.json", mode="r", encoding="utf-8") as f:
                try:
                    self.category = json.load(f)
                except:
                    self.category = {}
        # 第一次运行时创建forecast.cfg 用来储存self.forecast_funcs中新添加的项 以后则读取文件 储存在self.forecast_funcs中
        if not os.path.exists("configs/forecast.cfg"):
            with open("configs/forecast.cfg", mode="w", encoding="utf-8") as f:
                pass
        else:
            with open("configs/forecast.cfg", mode="r", encoding="utf-8") as f:
                forecast_lines = f.readlines()
                self.forecast_funcs = [forecast_line.strip() for forecast_line in forecast_lines]
        # 第一次运行时创建analyze.cfg 用来储存self.analyze_funcs中新添加的项 以后则读取文件 储存在self.analyze_funcs中
        if not os.path.exists("configs/analyze.cfg"):
            with open("configs/analyze.cfg", mode="w", encoding="utf-8") as f:
                pass
        else:
            with open("configs/analyze.cfg", mode="r", encoding="utf-8") as f:
                analyze_lines = f.readlines()
                self.analyze_funcs = [analyze_line.strip() for analyze_line in analyze_lines]
        # 第一次运行时创建plot.cfg 用来储存self.plot_funcs中新添加的项 以后则读取文件 储存在self.plot_funcs中
        if not os.path.exists("configs/plot.cfg"):
            with open("configs/plot.cfg", mode="w", encoding="utf-8") as f:
                pass
        else:
            with open("configs/plot.cfg", mode="r", encoding="utf-8") as f:
                plot_lines = f.readlines()
                self.plot_funcs = [plot_line.strip() for plot_line in plot_lines]
        # 第一次运行时创建templates.cfg 用来储存self.templates中新添加的项 以后则读取文件
        if not os.path.exists("configs/templates.cfg"):
            with open("configs/templates.cfg", mode="w", encoding="utf-8") as f:
                pass
        else:
            with open("configs/templates.cfg", mode="r", encoding="utf-8") as f:
                templates_lines = f.readlines()
                self.templates = [templates_line.strip() for templates_line in templates_lines]

    def update_file(self):
        # 更新year.cfg
        with open("configs/year.cfg", mode="w", encoding="utf-8") as f:
            for year in self.year:
                f.writelines(year + '\n')
        # 更新region.cfg
        with open("configs/region.cfg", mode="w", encoding="utf-8") as f:
            for region in self.region:
                f.writelines(region + '\n')
        # 更新region.cfg
        with open("configs/category.json", mode="w", encoding="utf-8") as f:
            temp_json = json.dumps(self.category, sort_keys=False, indent=4, separators=(',', ': '))
            f.write(temp_json)

    def update_func_file(self, path, funcs):
        with open(path, mode="w", encoding="utf-8") as f:
            for func in funcs:
                f.writelines(func + '\n')

    def update_templates(self):
        with open("configs/templates.cfg", mode='w', encoding='utf-8') as f:
            for template in self.templates:
                f.writelines(template + '\n')

    # 初始化gridLayout 向其添加checkBox
    def initialize_gridLayout(self, glo, array):
        positions = []
        # 创建year中的项对应的位置列表
        for i in range(len(array)):
            positions.append((i//3, i%3))
            # positions.append((i//4, i%4))
        # 创建checkBox并通过addWidget方法添加到布局中
        for position, name in zip(positions, array):
            checkBox = QCheckBox(name)
            # try to fix it (whh)
            checkBox.setStyleSheet("QCheckBox { width: 100px; height: 25px;}")  # 设置大小
            checkBox.setFont(QtGui.QFont("Microsoft YaHei", 10))  # 设置字体样式 字号
            checkBox.setCursor(Qt.PointingHandCursor)  # 设置鼠标图标
            checkBox.setFixedWidth(100)
            checkBox.setFixedHeight(30)
            glo.addWidget(checkBox, *position)

    def initialize_template_list(self):
        for item in self.templates:
            self.template_list.addItem(item)

    def clear_gridLayout(self, glo):
        gridLayout_widgets = self.get_gridLayout_widgets(glo)
        for widget in gridLayout_widgets:
            glo.removeWidget(widget)
            sip.delete(widget)

    # 获取gridLayout里的组件
    def get_gridLayout_widgets(self, glo):
        temp = []
        for i in range(glo.count()):
            temp.append(glo.itemAt(i).widget())
        return temp

    def initialize_listWidget(self, lsw, funcs):
        for func in funcs:
            lsw.addItem(func)

    # 初始化treeWidget 使其显示勾选框
    # add str for ctg (whh)
    def initiate_treeWidget(self, trw):
        for ctg1, val1 in self.category.items():
            child1 = QTreeWidgetItem(trw)
            child1.setText(0, str(ctg1))
            child1.setFlags(child1.flags() | Qt.ItemIsTristate)
            for ctg2, val2 in val1.items():
                child2 = QTreeWidgetItem(child1)
                child2.setText(0, str(ctg2))
                child2.setFlags(child2.flags() | Qt.ItemIsTristate)
                for ctg3, val3 in val2.items():
                    child3 = QTreeWidgetItem(child2)
                    child3.setText(0, str(ctg3))
                    child3.setFlags(child3.flags() | Qt.ItemIsTristate)
                    for ctg4, val4 in val3.items():
                        child4 = QTreeWidgetItem(child3)
                        child4.setText(0, str(ctg4))
                        child4.setFlags(child4.flags() | Qt.ItemIsTristate)
                        for ctg5, val5 in val4.items():
                            child5 = QTreeWidgetItem(child4)
                            child5.setText(0, str(ctg5))
                            child5.setFlags(child5.flags() | Qt.ItemIsTristate)
                            for ctg6, val6 in val5.items():
                                child6 = QTreeWidgetItem(child5)
                                child6.setText(0, str(ctg6))
                                child6.setCheckState(0, Qt.Unchecked)


    def findSelected(self, root, cat):
        if root.checkState(0) == Qt.Checked:
            cat.append(root.text(0))
            item = tbw_item()
            item.setName(cat)
            item.index = self.get_treeWidget_selected(root)
            self.contain.append(item)
            return

        if root.checkState(0) == Qt.PartiallyChecked:
            cat.append(root.text(0))

            child_cnt = root.childCount()
            for i in range(child_cnt):
                child_item = root.child(i)
                tmp = []
                for i in range(len(cat)):
                    tmp.append(cat[i])
                self.findSelected(child_item, tmp)


    def get_treeWidget_selected(self, trw):
        """
        递归获取treewidget所选的路径
        :param trw: 根节点
        """
        temp = []
        self.get_trw_selected_iterative(trw, temp)
        return temp


    def get_trw_selected_iterative(self, leveln_item, category):
        """
        get_treeWidget_selected的工具函数

        :param leveln_item: 当前搜索节点
        :param category: 存放当前搜索结果的列表
        """
        leveln_child_count = leveln_item.childCount()
        if leveln_child_count == 0:
            if leveln_item.checkState(0) == Qt.Checked:
                category.append(leveln_item.text(0).split(' ')[0])
            return
        else:
            for leveln_child_i in range(leveln_child_count):
                leveln_child_item = leveln_item.child(leveln_child_i)
                self.get_trw_selected_iterative(leveln_child_item, category)


    def set_trw_checkState_iterative(self, leveln_item, checked):
        """
        递归设置treewidget勾选

        :param leveln_item: 当前节点
        :param checked: 勾选状态
        """
        leveln_child_count = leveln_item.childCount()
        if leveln_child_count == 0:
            if checked:
                leveln_item.setCheckState(0, Qt.Unchecked)
            else:
                leveln_item.setCheckState(0, Qt.Checked)
            return
        else:
            for leveln_child_i in range(leveln_child_count):
                leveln_child_item = leveln_item.child(leveln_child_i)
                self.set_trw_checkState_iterative(leveln_child_item, checked)


    def form_region_from_db(self):
        """
        从数据库中读取所有地域
        """
        query = 'SELECT `地域` from `Data`'
        data = self.database.get_data_simple(query)
        self.region = []
        if data:
            for item in data:
                if item['地域'] not in self.region:
                    self.region.append(item['地域'])


    def form_year_from_db(self):
        """
        从数据库中读取所有年份
        """
        query = 'SELECT `年份` from `Data`'
        try:
            data = self.database.get_data_simple(query)
        except EXECUTE_FAILURE as exe_f:
            QMessageBox.critical(self, '错误', exef.__str__())
            return
        self.year = []
        if data:
            for item in data:
                years = json.loads(item['年份'])
                for key in years.keys():
                    if key not in self.year:
                        self.year.append(key)


    def form_category_structure_from_db(self):
        """
        从数据库中读取所有路径
        """
        query = 'SELECT * FROM `Data`'
        data = self.database.get_data(query)
        self.category = {}
        if data:
            for item in data:
                index = str(item['编码'])
                region = item['地域']
                ctg1 = item['级别1']
                ctg2 = item['级别2']
                ctg3 = item['级别3']
                ctg4 = item['级别4']
                ctg5 = item['级别5']
                ctg6 = item['级别6']
                final_name = '%s %s' % (index, ctg6)
                if ctg1 not in self.category.keys():
                    self.category[ctg1] = {ctg2: {ctg3: {ctg4: {ctg5: {final_name: [region]}}}}}
                else:
                    if ctg2 not in self.category[ctg1].keys():
                        self.category[ctg1][ctg2] = {ctg3: {ctg4: {ctg5: {final_name: [region]}}}}
                    else:
                        if ctg3 not in self.category[ctg1][ctg2].keys():
                            self.category[ctg1][ctg2][ctg3] = {ctg4: {ctg5: {final_name: [region]}}}
                        else:
                            if ctg4 not in self.category[ctg1][ctg2][ctg3].keys():
                                self.category[ctg1][ctg2][ctg3][ctg4] = {ctg5: {final_name: [region]}}
                            else:
                                if ctg5 not in self.category[ctg1][ctg2][ctg3][ctg4].keys():
                                    self.category[ctg1][ctg2][ctg3][ctg4][ctg5] = {final_name: [region]}
                                else:
                                    if final_name not in self.category[ctg1][ctg2][ctg3][ctg4][ctg5]:
                                        self.category[ctg1][ctg2][ctg3][ctg4][ctg5][final_name] = [region]
                                    elif region not in self.category[ctg1][ctg2][ctg3][ctg4][ctg5][final_name]:
                                        self.category[ctg1][ctg2][ctg3][ctg4][ctg5][final_name].append(region)



    def form_category_structure(self, rows):
        """
        递归设置treewidget内容
        :param rows: treewidget当前行数
        """
        # 数据从excel的第2行开始
        for row in rows:
            index = str(int(row[0]))
            region = row[1]
            ctg1 = row[2]
            ctg2 = row[3]
            ctg3 = row[4]
            ctg4 = row[5]
            ctg5 = row[6]
            ctg6 = row[7]
            final_name = '%s %s' % (index, ctg6)
            if ctg1 not in self.category.keys():
                self.category[ctg1] = {ctg2: {ctg3: {ctg4: {ctg5: {final_name: [region]}}}}}
            else:
                if ctg2 not in self.category[ctg1].keys():
                    self.category[ctg1][ctg2] = {ctg3: {ctg4: {ctg5: {final_name: [region]}}}}
                else:
                    if ctg3 not in self.category[ctg1][ctg2].keys():
                        self.category[ctg1][ctg2][ctg3] = {ctg4: {ctg5: {final_name: [region]}}}
                    else:
                        if ctg4 not in self.category[ctg1][ctg2][ctg3].keys():
                            self.category[ctg1][ctg2][ctg3][ctg4] = {ctg5: {final_name: [region]}}
                        else:
                            if ctg5 not in self.category[ctg1][ctg2][ctg3][ctg4].keys():
                                self.category[ctg1][ctg2][ctg3][ctg4][ctg5] = {final_name: [region]}
                            else:
                                if final_name not in self.category[ctg1][ctg2][ctg3][ctg4][ctg5]:
                                    self.category[ctg1][ctg2][ctg3][ctg4][ctg5][final_name] = [region]
                                elif region not in self.category[ctg1][ctg2][ctg3][ctg4][ctg5][final_name]:
                                    self.category[ctg1][ctg2][ctg3][ctg4][ctg5][final_name].append(region)

    def is_category_exist(self, index, region):
        for ctg1, val1 in self.category.items():
            for ctg2, val2 in val1.items():
                for ctg3, val3 in val2.items():
                    for ctg4, val4 in val3.items():
                        for ctg5, val5 in val4.items():
                            for final_name, regions in val5.items():
                                i = final_name.split(' ')[0]
                                if index == i and region in regions:
                                    return True, ctg1, ctg2, ctg3, ctg4, ctg5, final_name
        return False, None, None, None, None, None, None


    def delete_category(self, region, ctg1, ctg2, ctg3, ctg4, ctg5, final_name):
        """
        递归删除treewidget中某路径

        :param region: 数据项的地域
        :param ctg1, ctg2, ctg3, ctg4, ctg5, ctg6: 分别对应级别1到6
        :param final_name: 最后一个数据项
        """
        try:
            self.category[ctg1][ctg2][ctg3][ctg4][ctg5][final_name].remove(region)
        except:
            return
        # 删除完之后列表为空
        if not self.category[ctg1][ctg2][ctg3][ctg4][ctg5][final_name]:
            self.category[ctg1][ctg2][ctg3][ctg4][ctg5].pop(final_name)
        # 删除完之后final_name对应的项后ctg5字典为空
        if not self.category[ctg1][ctg2][ctg3][ctg4][ctg5]:
            self.category[ctg1][ctg2][ctg3][ctg4].pop(ctg5)
        # 删除完之后ctg5对应的项后ctg4字典为空
        if not self.category[ctg1][ctg2][ctg3][ctg4]:
            self.category[ctg1][ctg2][ctg3].pop(ctg4)
        # 删除完之后ctg4对应的项后ctg3字典为空
        if not self.category[ctg1][ctg2][ctg3]:
            self.category[ctg1][ctg2].pop(ctg3)
        # 删除完之后ctg3对应的项后ctg2字典为空
        if not self.category[ctg1][ctg2]:
            self.category[ctg1].pop(ctg2)
        # 删除完之后ctg2对应的项后ctg1字典为空
        if not self.category[ctg1]:
            self.category.pop(ctg1)


    def add_data_from_excel(self, filepath):
        """
        从excel表格中读取数据

        :param filepath: excel表格的文件路径
        """
        workbook = xlrd.open_workbook(filepath)
        worksheet = workbook.sheet_by_index(0)
        header = worksheet.row_values(0)  # 读取excel第1行的header
        if header[0] == '编码':
            regions = worksheet.col_values(1)
            year_list = []
            for year_index in range(8, len(header)):
                # 遇到预测数据就终止
                try:
                    year = str(int(header[year_index]))
                except:
                    break
                year_list.append(year)

            try:
                self.database.add_data_excel(worksheet, year_list)
            except INSERT_FAILURE as insrtf:
                QMessageBox.critical(self, "错误", insrtf.__str__())
                return

            for year in year_list:
                if year not in self.year and year:
                    self.year.append(year)
                    # sort the year (whh)
                    self.year.sort()

            for region in regions[1:]:
                if region not in self.region and region:
                    self.region.append(region)
            rows = []
            for row_num in range(1, worksheet.nrows):
                rows.append(worksheet.row_values(row_num))
            self.form_category_structure(rows)

            self.update_file()  # 更新year.cfg region.cfg
            self.clear_gridLayout(self.glo_city_search)  # 清空gridLayout
            self.clear_gridLayout(self.glo_year_search)  # 清空gridLayout
            self.clear_gridLayout(self.glo_city_process)  # 清空gridLayout
            self.clear_gridLayout(self.glo_year_process)  # 清空gridLayout
            self.clear_gridLayout(self.glo_city_draw)  # 清空gridLayout
            self.clear_gridLayout(self.glo_year_draw)  # 清空gridLayout
            self.initialize_gridLayout(self.glo_city_search, self.region)  # 按照新的self.region更新gridLayout
            self.initialize_gridLayout(self.glo_year_search, self.year)  # 按照新的self.year更新gridLayout
            self.initialize_gridLayout(self.glo_city_process, self.region)  # 按照新的self.region更新gridLayout
            self.initialize_gridLayout(self.glo_year_process, self.year)  # 按照新的self.year更新gridLayout

            self.initialize_gridLayout(self.glo_city_draw, self.region)  # 按照新的self.region更新gridLayout
            self.initialize_gridLayout(self.glo_year_draw, self.year)  # 按照新的self.year更新gridLayout

            self.trw_search.clear()
            self.trw_process.clear()
            self.trw_process_2.clear()
            self.initiate_treeWidget(self.trw_search)
            self.initiate_treeWidget(self.trw_process)
            self.initiate_treeWidget(self.trw_process_2)


    def load_forecast_func(self):
        """
        读取预测函数文件
        """
        filepath = QtWidgets.QFileDialog.getOpenFileName(self, "打开Python文件", "", "Python文件 (*.py)")
        if filepath[0] != '':
            filename = filepath[0].split('/')[-1].split('.')[0]
            if filename not in self.forecast_funcs:
                self.forecast_funcs.append(filename)
                self.update_func_file("configs/forecast.cfg", self.forecast_funcs)
                self.lsw_forecast.addItem(filename)
            dst_path = os.path.join("funcs/forecast_funcs", filepath[0].split('/')[-1])
            shutil.copy(filepath[0], dst_path)

    def load_analyze_func(self):
        """
        读取分析函数文件
        """
        filepath = QtWidgets.QFileDialog.getOpenFileName(self, "打开Python文件", "", "Python文件 (*.py)")
        if filepath[0] != '':
            filename = filepath[0].split('/')[-1].split('.')[0]
            if filename not in self.analyze_funcs:
                self.analyze_funcs.append(filename)
                self.update_func_file("configs/analyze.cfg", self.analyze_funcs)
                self.lsw_analyze.addItem(filename)
            dst_path = os.path.join("funcs/analyze_funcs", filepath[0].split('/')[-1])
            shutil.copy(filepath[0], dst_path)


    def load_plot_func(self):
        """
        读取绘图函数文件
        """
        filepath = QtWidgets.QFileDialog.getOpenFileName(self, "打开Python文件", "", "Python文件 (*.py)")
        if filepath[0] != '':
            filename = filepath[0].split('/')[-1].split('.')[0]
            if filename not in self.plot_funcs:
                self.plot_funcs.append(filename)
                self.update_func_file("configs/plot.cfg", self.plot_funcs)
                self.lsw_plot_2.addItem(filename)
            dst_path = os.path.join("funcs/plot_funcs", filepath[0].split('/')[-1])
            shutil.copy(filepath[0], dst_path)
    

    def load_templates(self):
        """
        读取模板文件
        """
        filepath = QtWidgets.QFileDialog.getOpenFileName(self, "打开Excel文件", "", "Excel文件 (*.xls *xlsx)")
        if filepath[0] != '':
            filename = filepath[0].split('/')[-1].split('.')[0]
            if filename not in self.templates:
                self.templates.append(filename)
                self.update_templates()
                self.template_list.addItem(filename)
            dst_path = os.path.join("templates", filepath[0].split('/')[-1])
            shutil.copy(filepath[0], dst_path)
    
   
    def buttonForRow(self, id):
        """
        设置查询结果表格中每一行最后的两个按钮

        :param id: 表格行号
        """
        widget = QWidget()

        detailBtn = QPushButton("查看详情")
        detailBtn.setStyleSheet("color: rgb(255, 255, 255); background-color: rgb(75, 225, 75); font-family:'微软雅黑'; font-Point Size:10")
        detailBtn.clicked.connect(lambda:self.view_detail(id))
        # self.detail_for_row.append(detailBtn)

        modifyBtn = QPushButton("修正")
        modifyBtn.setStyleSheet("color: rgb(255, 255, 255); background-color: rgb(225, 174, 0); font-family:'微软雅黑'; font-Point Size:10")
        # self.modify_for_row.append(modifyBtn)
        modifyBtn.clicked.connect(lambda:self.modify(id, 1))
        hLayout = QHBoxLayout()
        hLayout.addWidget(detailBtn)
        hLayout.addWidget(modifyBtn)
        hLayout.setContentsMargins(5,2,5,2)
        widget.setLayout(hLayout)
        return widget


    def button_for_row(self, id):
        widget = QWidget()

        detailBtn = QPushButton("添加相关")
        detailBtn.setStyleSheet("color: rgb(255, 255, 255); background-color: rgb(225, 174, 0); font-family:'微软雅黑'; font-Point Size:10")
        detailBtn.clicked.connect(lambda:self.on_btn_add_in_result_dialog(id))
        # detailBtn.clicked.connect(lambda:self.view_detail(id))
        # self.detail_for_row.append(detailBtn)

        modifyBtn = QPushButton("修正相关")
        modifyBtn.setStyleSheet("color: rgb(255, 255, 255); background-color: rgb(75, 225, 75); font-family:'微软雅黑'; font-Point Size:10")
        modifyBtn.clicked.connect(lambda:self.on_btn_modify_in_result_dialog(id))
        # self.modify_for_row.append(modifyBtn)
        # modifyBtn.clicked.connect(lambda:self.modify(id))
        hLayout = QHBoxLayout()
        hLayout.addWidget(detailBtn)
        hLayout.addWidget(modifyBtn)
        hLayout.setContentsMargins(5,2,5,2)
        widget.setLayout(hLayout)
        return widget


    def search(self):
        """
        主窗口的搜索功能
        """
        goal = self.le_search.text()
        if goal != '':
            self.search_path = []
            roots = []
            for i in range(self.trw_search.invisibleRootItem().childCount()):
                roots.append(self.trw_search.invisibleRootItem().child(i))

    
            for root in roots:
                tmp = []
                if self.is_exist(root, goal):
                    self.exact_find_goal(root, goal, tmp, False)
                else:
                    self.fuzzy_find_goal(root, goal, tmp, False)
            
            if self.search_path:
                self.dialog_result = dialog_search_result()
                self.dialog_result.setupTbw(self.search_path)
                for i in range(self.dialog_result.tbw_result.rowCount()):
                    self.dialog_result.tbw_result.setCellWidget(i, 1, self.button_for_row(i))
                self.dialog_result.exec_()

            else:
                QMessageBox.warning(self, "警告", "未找到相关路径")


    def is_exist(self, root, goal):
        """
        search方法的工具函数，判断所找的路径是否存在

        :param root: 搜索起点节点
        :param goal: 搜索目标
        """
        if self.exact_match(root.text(0), goal):
            return True
        if root.childCount() == 0 and not self.exact_match(root.text(0), goal):
            return False
        b = False
        for i in range(root.childCount()):
            b = b or self.is_exist(root.child(i), goal)
        return b 

 
    def exact_find_goal(self, root, goal, path, find):
        """
        递归精确查找路径（服务于search函数）

        :param root: 搜索起点节点
        :param goal: 搜索目标
        :param path: 搜到的路径
        :param find: 是否找到
        """
        path.append(root.text(0))
        if self.exact_match(root.text(0), goal):
        # if root.text(0) == goal:
            find = True
           
        child_cnt = root.childCount()
        if child_cnt == 0 and find == True:
            self.search_path.append(path)
            return
        for i in range(child_cnt):
            tmp = []
            for j in range(len(path)):
                tmp.append(path[j])
            child_item = root.child(i)
            self.exact_find_goal(child_item, goal, tmp, find)


    def fuzzy_find_goal(self, root, goal, path, find):
        """
        递归模糊查找相关路径（服务于search函数）

        :param root: 搜索起点节点
        :param goal: 搜索目标
        :param path: 搜到的路径
        :param find: 是否找到
        """
        path.append(root.text(0))
        if self.fuzzy_match(root.text(0), goal):
            find = True
            
        child_cnt = root.childCount()
        if child_cnt == 0 and find == True:
            self.search_path.append(path)
            return
        for i in range(child_cnt):
            tmp = []
            for j in range(len(path)):
                tmp.append(path[j])
            child_item = root.child(i)
            self.fuzzy_find_goal(child_item, goal, tmp, find)


    def exact_match(self, key, goal):
        cnt = 0
        for i in goal:
            for j in key:
                if i == j:
                    cnt += 1
                    break

        if cnt / len(key) >= 0.8:
            return True
        else:
            return False


    #模糊匹配的方法（服务于find_goal函数）
    def fuzzy_match(self, key, goal):
        for i in goal:
            for j in key:
                if i == j:
                    return True
        return False

# ======================================================自定义方法======================================================

# ======================================================UI组件方法======================================================
    # 查询更新-地域 地域checkBox全选
    @QtCore.pyqtSlot()
    def on_btn_search_region_select_all_clicked(self):
        region_checkboxes = self.get_gridLayout_widgets(self.glo_city_search)
        for checkbox in region_checkboxes:
            checkbox.setChecked(not self.city_search_checked)
        self.city_search_checked = not self.city_search_checked

    # 查询更新-年份 年份checkBox全选
    @QtCore.pyqtSlot()
    def on_btn_search_year_select_all_clicked(self):
        year_checkboxes = self.get_gridLayout_widgets(self.glo_year_search)
        for checkbox in year_checkboxes:
            checkbox.setChecked(not self.year_search_checked)
        self.year_search_checked = not self.year_search_checked

    # 查询更新-指标 指标checkBox全选
    @QtCore.pyqtSlot()
    def on_btn_search_category_select_all_clicked(self):
        root = self.trw_search.invisibleRootItem()
        self.set_trw_checkState_iterative(root, self.category_search_checked)
        self.category_search_checked = not self.category_search_checked

    # 查询更新-增添 用户增添新的数据
    @QtCore.pyqtSlot()
    def on_btn_search_add_clicked(self):
        if self.region:
            header = ['编码', '地域', '级别1', '级别2', '级别3', '级别4', '级别5', '级别6']
            for year in self.year:
                header.append(year)
            dialog_1 = Dialog_1(0, header)
            if dialog_1.exec_() == 1:
                data = dialog_1.get_data1()

                try:
                    self.database.add_data_dialog(data, self.year)
                except INSERT_FAILURE as insrtf:
                    QMessageBox.critical(self, "错误", insrtf.__str__())
                    return

                index = data[0]
                region = data[1]
                is_exist, old_ctg1, old_ctg2, old_ctg3, old_ctg4, old_ctg5, old_final_name = self.is_category_exist(
                    index, region)

                if is_exist:
                    self.delete_category(region, old_ctg1, old_ctg2, old_ctg3, old_ctg4, old_ctg5, old_final_name)

                if data[1] not in self.region:
                    self.region.append(data[1])
                    self.clear_gridLayout(self.glo_city_search)  # 清空gridLayout
                    self.clear_gridLayout(self.glo_city_process)  # 清空gridLayout
                    self.clear_gridLayout(self.glo_city_draw)  # 清空gridLayout
         
                    self.initialize_gridLayout(self.glo_city_search, self.region)  # 按照新的self.region更新gridLayout
                    self.initialize_gridLayout(self.glo_city_process, self.region)  # 按照新的self.region更新gridLayout
                    self.initialize_gridLayout(self.glo_city_draw, self.region)  # 按照新的self.region更新gridLayout
           
                self.form_category_structure([data])
                self.update_file()  # 更新year.cfg region.cfg

                self.trw_search.clear()
                self.trw_process.clear()
                self.trw_process_2.clear()
                self.initiate_treeWidget(self.trw_search)
                self.initiate_treeWidget(self.trw_process)
                self.initiate_treeWidget(self.trw_process_2)
               

                cat = tbw_item()
                cat.index.append(index)
                cat.year = self.year
                cat.row = len(self.contain_tbl)
                cat_data = []
                dic = {}
                for i in range(len(data)):
                    dic[header[i]] = data[i]
                cat_data.append(dic)
                cat.data = cat_data
                cat.name = data[2 : 8]
                cat.setRealName()

                self.contain_tbl.append(cat)

                self.reload_table_widget()


    def on_btn_modify_in_result_dialog(self, id):
       
        # print(id)
        path = self.search_path[id]
        header = ['编码', '地域', '级别1', '级别2', '级别3', '级别4', '级别5', '级别6']
        for year in self.year:
            header.append(year)
        index = path[len(path) - 1].split(' ')[0]
        search_modify_dialog = search_modify(header, index, self.region)
        if search_modify_dialog.exec_() == 1:
            data = search_modify_dialog.data

            modified_data = search_modify_dialog.get_data()

            modified_place = []
            for i in range(2, len(data)):
                if str(data[i]) != str(modified_data[i]):
                    modified_place.append((0, i))

            index = modified_data[0]
            region = modified_data[1]
            is_exist, old_ctg1, old_ctg2, old_ctg3, old_ctg4, old_ctg5, old_final_name = self.is_category_exist(
                index, region)

            # data = search_modify_dialog.data
            modd = []
            modd.append(modified_data)

            try:
                self.database.modify_data(modified_place, modd, header)
            except UPDATE_FAILURE as updf:
                QMessageBox.critical(self, "错误", updf.__str__())
                return

            try:
                self.dialog_result.modify_data(id, modified_data, 0)
            except UPDATE_FAILURE as updf:
                QMessageBox.critical(self, "错误", updf.__str__())

            self.delete_category(region, data[2], data[3], data[4], data[5], data[6], '%s %s' % (data[0], data[7]))
            
            
            if is_exist:
                self.delete_category(region, old_ctg1, old_ctg2, old_ctg3, old_ctg4, old_ctg5, old_final_name)
            
            self.form_category_structure([modified_data])
           
            if modified_data[1] not in self.region:
                self.region.append(modified_data[1])
                self.clear_gridLayout(self.glo_city_search)  # 清空gridLayout
                self.clear_gridLayout(self.glo_city_process)  # 清空gridLayout
                self.clear_gridLayout(self.glo_city_draw)  # 清空gridLayout
            
                self.initialize_gridLayout(self.glo_city_search, self.region)  # 按照新的self.region更新gridLayout
                self.initialize_gridLayout(self.glo_city_process, self.region)  # 按照新的self.region更新gridLayout
                self.initialize_gridLayout(self.glo_city_draw, self.region)  # 按照新的self.region更新gridLayout
            
            self.update_file()  # 更新year.cfg region.cfg

            self.trw_search.clear()
            self.trw_process.clear()
            self.trw_process_2.clear()
            self.initiate_treeWidget(self.trw_search)
            self.initiate_treeWidget(self.trw_process)
            self.initiate_treeWidget(self.trw_process_2)


    def on_btn_add_in_result_dialog(self, id):
        # print(id)
        header = ['编码', '地域', '级别1', '级别2', '级别3', '级别4', '级别5', '级别6']
        for year in self.year:
            header.append(year)
        
        data = []
        path = self.search_path[id]
        for i in range(len(path) - 1):
            data.append(path[i])
        data.append(path[len(path) - 1].split(' ')[1])
        data.append(path[len(path) - 1].split(' ')[0])

        search_add_dialog = search_add(header, data)
        if search_add_dialog.exec_() == 1:
            data_added = search_add_dialog.get_data()
            index = data_added[0]
            region = data_added[1]

            try:
                self.database.add_data_dialog(data_added, self.year)
                self.dialog_result.modify_data(id, data_added, 1)
            except INSERT_FAILURE as insrtf:
                QMessageBox.critical(search_add_dialog, "错误", insrtf.__str__())
                return

            is_exist, old_ctg1, old_ctg2, old_ctg3, old_ctg4, old_ctg5, old_final_name = self.is_category_exist(
                    index, region)

            if is_exist:
                    self.delete_category(region, old_ctg1, old_ctg2, old_ctg3, old_ctg4, old_ctg5, old_final_name)
            if region not in self.region:
                self.region.append(region)
                self.clear_gridLayout(self.glo_city_search)  # 清空gridLayout
                self.clear_gridLayout(self.glo_city_process)  # 清空gridLayout
                self.clear_gridLayout(self.glo_city_draw)  # 清空gridLayout
            
                self.initialize_gridLayout(self.glo_city_search, self.region)  # 按照新的self.region更新gridLayout
                self.initialize_gridLayout(self.glo_city_process, self.region)  # 按照新的self.region更新gridLayout
                self.initialize_gridLayout(self.glo_city_draw, self.region)  # 按照新的self.region更新gridLayout
            
            
            self.form_category_structure([data_added])
            self.update_file()  # 更新year.cfg region.cfg

            self.trw_search.clear()
            self.trw_process.clear()
            self.trw_process_2.clear()
            self.initiate_treeWidget(self.trw_search)
            self.initiate_treeWidget(self.trw_process)
            self.initiate_treeWidget(self.trw_process_2)
          

    def view_detail(self, id):
        """
        查询结果表格中每一行的显示详情按钮逻辑
        """
        self.dialog_detail = detail_dialog(id = id, item = self.tbw_search_class.datas[id])
        self.dialog_detail.btnModi.clicked.connect(lambda:self.modify(id, 0))
        name = self.tbw_search_class.datas[id].realname
        self.dialog_detail.setTitle(name)
        self.dialog_detail.table_display()
        self.dialog_detail.exec_()
    

    def modify(self, id, mode):
        """
        查询结果表格中每一行的修改按钮逻辑
        """
        item = self.tbw_search_class.datas[id]
        header = ['编码', '地域', '级别1', '级别2', '级别3', '级别4', '级别5', '级别6']
        for year in item.year:
            header.append(year)
        
        data = []
        region = item.data[0]['地域']
        
        for i in item.index:
            condition_str = '`编码`=' + "%d and " % int(i) + '`地域`=' + "'%s'" % region
            query = 'select * from Data where ' + condition_str
            # print("sql语句")
            # print(query)
            try:
                result = self.database.get_data(query)[0]
            except EXECUTE_FAILURE as exef:
                QMessageBox.critical(self, '错误', exef.__str__())
                return

            tmp = []
            for i in header:
                if i in result.keys():
                    tmp.append(result[i])
                else:
                    header.remove(i)
                    continue

            data.append(tmp)
           
        dialog_1 = Dialog_1(1, header, data=data)
        if dialog_1.exec_() == 1:
            modified_data = dialog_1.get_data()  #修改后的表格中的数据
            modified_place = []
            for i in range(len(data)):   #对比修改前和修改后的值，将有修改过的值的坐标存入modified_data中
                for j in range(2, len(data[0])):
                    if data[i][j] != modified_data[i][j]:
                        modified_place.append((i, j))
            try:
                self.database.modify_data(modified_place, modified_data, header)  #更新数据库
            except UPDATE_FAILURE as updf:
                QMessageBox.critical(self, "错误", updf.__str__())
                return
            
            for place in modified_place:
                row = place[0]; col = place[1]
                item.data[row][header[col]] = modified_data[row][col]

            if mode == 0:
                self.dialog_detail.table_display()
            
           
            for i in range(len(data)):
                self.reload_after_modify(data[i], modified_data[i], item.year)


    def reload_after_modify(self, data, modified_data, year):
        """
        修改后重新加载主窗口（"确定"按钮按下后调用）
        """
        # print("reload_after_modify in ui_mainwindow")
        index = modified_data[0]
        region = modified_data[1]
        is_exist, old_ctg1, old_ctg2, old_ctg3, old_ctg4, old_ctg5, old_final_name = self.is_category_exist(
            index, region)
        
        self.delete_category(region, data[2], data[3], data[4], data[5], data[6], '%s %s' % (str(data[0]), data[7]))
        
        if is_exist:
            self.delete_category(region, old_ctg1, old_ctg2, old_ctg3, old_ctg4, old_ctg5, old_final_name)

        self.form_category_structure([modified_data])

        if modified_data[1] not in self.region:
            self.region.append(modified_data[1])
            self.clear_gridLayout(self.glo_city_search)  # 清空gridLayout
            self.clear_gridLayout(self.glo_city_process)  # 清空gridLayout
            self.clear_gridLayout(self.glo_city_draw)  # 清空gridLayout
    
            self.initialize_gridLayout(self.glo_city_search, self.region)  # 按照新的self.region更新gridLayout
            self.initialize_gridLayout(self.glo_city_process, self.region)  # 按照新的self.region更新gridLayout
            self.initialize_gridLayout(self.glo_city_draw, self.region)  # 按照新的self.region更新gridLayout
            
        self.update_file()  # 更新year.cfg region.cfg

        self.trw_search.clear()
        self.trw_process.clear()
        self.trw_process_2.clear()
        self.initiate_treeWidget(self.trw_search)
        self.initiate_treeWidget(self.trw_process)
        self.initiate_treeWidget(self.trw_process_2)


    # 查询更新-更新 导入excel
    @QtCore.pyqtSlot()
    def on_btn_search_update_clicked(self):
        # todo
        filename = QtWidgets.QFileDialog.getOpenFileName(self, "打开Excel文件", "", "Excel文件 (*.xls *xlsx)")
        if filename[0] != '':
            self.add_data_from_excel(filename[0])


    # 查询更新-导出 导出excel
    @QtCore.pyqtSlot()
    def on_btn_search_export_clicked(self):
        if self.tbw_search_class.datas:
            date = time.strftime("%Y_%m_%d_%H-%M-%S", time.localtime())
            # try to fix it (whh)
            # filename = QtWidgets.QFileDialog.getSaveFileName(self, "导出Excel", date, "Excel files (*.xlsx)")
            filename = QtWidgets.QFileDialog.getSaveFileName(self, "导出Excel", date, "Excel files (*.xls)")
            if filename[0] != '':
                self.tbw_search_class.export_excel(filename[0])


    # 查询更新-清空 清空tableWidget
    @QtCore.pyqtSlot()
    def on_btn_search_clear_clicked(self):
        self.tbw_search_class.clear()


    # 查询更新-清空数据库
    @QtCore.pyqtSlot()
    def on_btn_search_remove_all_clicked(self):
        reply = QMessageBox.question(self, '清空数据库', '您确定要清空数据库吗?\n注意：清空数据库后会自动退出程序！',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                self.database.empty_table()
            except EXECUTE_FAILURE as exef:
                QMessageBox.critical(self, '错误', exef.__str__())
                return
            
            for root, dirs, files in os.walk(r'configs'):
                os.remove(os.path.join(root, "category.json"))
                os.remove(os.path.join(root, "region.cfg"))
                os.remove(os.path.join(root, "year.cfg"))
            # 退出程序
            # 得到一个实例
            app = QApplication.instance()
            # 退出应用程序
            app.quit()


    def query_for_draw(self):
        self.draw_class.datas = []
        self.selected_year_draw = []  # 存放当前在gridLayout中被选中的年份
        self.selected_region_draw = []  # 存放当前在gridLayout中被选中的地域
        root = self.trw_process_2.invisibleRootItem()
        self.selected_category_draw = self.get_treeWidget_selected(root)  # 存放当前在tableWidget中被选中的指标
        
        region_checkboxes = self.get_gridLayout_widgets(self.glo_city_draw)
        for checkbox in region_checkboxes:
            if checkbox.isChecked():
                self.selected_region_draw.append(checkbox.text())
        year_checkboxes = self.get_gridLayout_widgets(self.glo_year_draw)
        for checkbox in year_checkboxes:
            if checkbox.isChecked():
                self.selected_year_draw.append(checkbox.text())
        
        if not self.selected_region_draw or not self.selected_year_draw or not self.selected_category_draw:
            return

        temp_str = '`编码`,`地域`,'
        condition_str = '('
        for region_index in range(len(self.selected_region_draw)):
            if region_index != len(self.selected_region_draw) - 1:
                condition_str += '`地域`=' + "'%s' or " % self.selected_region_draw[region_index]
            else:
                condition_str += '`地域`=' + "'%s') and (" % self.selected_region_draw[region_index]
        for category_index in range(len(self.selected_category_draw)):
            if category_index != len(self.selected_category_draw) - 1:
                condition_str += '`编码`=' + "%d or " % int(self.selected_category_draw[category_index])
            else:
                condition_str += '`编码`=' + "%d)" % int(self.selected_category_draw[category_index])
        for cat_index in range(1, 7):
            temp_str += '`级别%d`,' % cat_index
        
        temp_str += '`年份`'
        query = 'select %s from `Data` where %s' % (temp_str, condition_str)

        try:
            datas = self.database.get_data(query)
        except EXECUTE_FAILURE as exef:
            QMessageBox.critical(self, "错误", exef.__str__())
            return

        year = self.year
        for item in year:
            if item not in self.selected_year_draw:
                for dic in datas:
                    if item in dic.keys():
                        dic.pop(item)

        self.draw_data = datas

    # 查询更新-复杂查询
    @QtCore.pyqtSlot()
    def on_btn_search_complex_clicked(self):
        """
        "查询"按钮的逻辑
        """
        self.selected_year_search = []  # 存放当前在gridLayout中被选中的年份
        self.selected_region_search = []  # 存放当前在gridLayout中被选中的地域
        root = self.trw_search.invisibleRootItem()
        # tmp = []
        self.contain.clear()
        self.contain_tbl.clear()
        child_cnt = root.childCount()
        for i in range(child_cnt):
            tmp = []
            self.findSelected(root.child(i), tmp)
      

        region_checkboxes = self.get_gridLayout_widgets(self.glo_city_search)
        for checkbox in region_checkboxes:
            if checkbox.isChecked():
                self.selected_region_search.append(checkbox.text())
        year_checkboxes = self.get_gridLayout_widgets(self.glo_year_search)
        for checkbox in year_checkboxes:
            if checkbox.isChecked():
                self.selected_year_search.append(checkbox.text())

        if not self.selected_region_search or not self.selected_year_search or not self.contain:
            return

        for region in self.selected_region_search:
            for item in self.contain:
                cat = tbw_item()
                cat.data = item.data
                cat.name = item.name
                cat.index = item.index
                cat.year = item.year
                cat.row = item.row
                cat.realname = item.realname

                temp_str = '`编码`,`地域`,'
                condition_str = '(' + '`地域`=' + "'%s') and (" % region
                for cat_index in range(len(cat.index)):
                    if cat_index != len(cat.index) - 1:
                        condition_str += '`编码`=' + "%d or " % int(cat.index[cat_index])
                    else:
                        condition_str += '`编码`=' + "%d)" % int(cat.index[cat_index])

                for id in range(1, 7):
                    temp_str += '`级别%d`,' % id
              
                temp_str += '`年份`'
                query = 'select %s from Data where %s' % (temp_str, condition_str)
                
                # print("sql语句：")
                # print(query)

                try:
                    datas = self.database.get_data(query)
                except EXECUTE_FAILURE as exef:
                    QMessageBox.critical(self, '错误', exef.__str__())
                    return

                if datas:
                    cat.setData(datas)
                    cat.year = self.selected_year_search
                    self.contain_tbl.append(cat)
    
        if self.contain_tbl:
            
            self.tbw_search_class.add_data(self.contain_tbl, self.selected_year_search)
            self.tbw_search_class.display()
            self.tbw_search_class.widget.clearSelection()
            # print('\n\n')
            for i in range(self.tbw_search_class.widget.rowCount()):
                self.tbw_search_class.widget.setCellWidget(i, 1, self.buttonForRow(i))
            

    def reload_table_widget(self):
        """
        修改数据后重新加载查询结果表格
        """
        self.tbw_search_class.datas = self.contain_tbl
        self.tbw_search_class.display()
        for i in range(self.tbw_search_class.widget.rowCount()):
            self.tbw_search_class.widget.setCellWidget(i, 1, self.buttonForRow(i))


    # ======================================================查询更新

    # ======================================================预测分析
    # 预测分析-地域 地域checkBox全选
    @QtCore.pyqtSlot()
    def on_btn_process_region_select_all_clicked(self):
        region_checkboxes = self.get_gridLayout_widgets(self.glo_city_process)
        for checkbox in region_checkboxes:
            checkbox.setChecked(not self.city_process_checked)
        self.city_process_checked = not self.city_process_checked
    
    @QtCore.pyqtSlot()
    def on_btn_process_region_select_all_2_clicked(self):
        region_checkboxes = self.get_gridLayout_widgets(self.glo_city_draw)
        for checkbox in region_checkboxes:
            checkbox.setChecked(not self.city_draw_checked)
        self.city_draw_checked = not self.city_draw_checked

    # 预测分析-年份 年份checkBox全选
    @QtCore.pyqtSlot()
    def on_btn_process_year_select_all_clicked(self):
        year_checkboxes = self.get_gridLayout_widgets(self.glo_year_process)
        for checkbox in year_checkboxes:
            checkbox.setChecked(not self.year_process_checked)
        self.year_process_checked = not self.year_process_checked

    @QtCore.pyqtSlot()
    def on_btn_process_year_select_all_2_clicked(self):
        year_checkboxes = self.get_gridLayout_widgets(self.glo_year_draw)
        for checkbox in year_checkboxes:
            checkbox.setChecked(not self.year_draw_checked)
        self.year_draw_checked = not self.year_draw_checked

    # 预测分析-指标 指标checkBox全选
    @QtCore.pyqtSlot()
    def on_btn_process_category_select_all_clicked(self):
        root = self.trw_process.invisibleRootItem()
        self.set_trw_checkState_iterative(root, self.category_process_checked)
        self.category_process_checked = not self.category_process_checked

    @QtCore.pyqtSlot()
    def on_btn_process_category_select_all_2_clicked(self):
        root = self.trw_process_2.invisibleRootItem()
        self.set_trw_checkState_iterative(root, self.category_draw_checked)
        self.category_draw_checked = not self.category_draw_checked

    # 预测分析-预测函数
    @QtCore.pyqtSlot()
    def on_btn_process_add_forecast_clicked(self):
        self.load_forecast_func()

    # 预测分析-预测
    @QtCore.pyqtSlot()
    def on_btn_process_forecast_clicked(self):
        self.selected_year_process = []  # 存放当前在gridLayout中被选中的年份
        self.selected_region_process = []  # 存放当前在gridLayout中被选中的地域
        root = self.trw_process.invisibleRootItem()
        self.selected_category_process = self.get_treeWidget_selected(root)  # 存放当前在tableWidget中被选中的指标
        region_checkboxes = self.get_gridLayout_widgets(self.glo_city_process)
        for checkbox in region_checkboxes:
            if checkbox.isChecked():
                self.selected_region_process.append(checkbox.text())
        year_checkboxes = self.get_gridLayout_widgets(self.glo_year_process)
        for checkbox in year_checkboxes:
            if checkbox.isChecked():
                self.selected_year_process.append(checkbox.text())

        if not self.selected_region_process or not self.selected_category_process:
            return
        # elif len(self.selected_year_process) < 5:
        #     QMessageBox.warning(self, '警告','请选择至少5个年份以保证预测的准确性。')
        #     return

        temp_str = '`编码`,`地域`,'
        condition_str = '('
        for region_index in range(len(self.selected_region_process)):
            if region_index != len(self.selected_region_process) - 1:
                condition_str += '`地域`=' + "'%s' or " % self.selected_region_process[region_index]
            else:
                condition_str += '`地域`=' + "'%s') and (" % self.selected_region_process[region_index]
        for category_index in range(len(self.selected_category_process)):
            if category_index != len(self.selected_category_process) - 1:
                condition_str += '`编码`=' + "%d or " % int(self.selected_category_process[category_index])
            else:
                condition_str += '`编码`=' + "%d)" % int(self.selected_category_process[category_index])
        for cat_index in range(1, 7):
            temp_str += '`级别%d`,' % cat_index
        
        temp_str += '`年份`'
       
        query = 'select %s from `Data` where %s' % (temp_str, condition_str)
        try:
            datas = self.database.get_data(query)
        except EXECUTE_FAILURE as exef:
            QMessageBox.critical(self, '错误', exef.__str__())
            return

        year = [year for year in range(0, 2020)]
        # for item in year:
        #     if item not in self.selected_year_process:
        #         for dic in datas:
        #             if item in dic.keys():
        #                 dic.pop(item)
                        
      
        if datas and self.lsw_forecast.selectedItems():

            #  生成输入到预测函数中的x numpy数组和y numpy数组
            years = [year for year in self.selected_year_process]
            not_in_key = []
            for item in years:
                if item not in datas[0].keys():
                    not_in_key.append(item)

            for item in not_in_key:
                years.remove(item)

            if len(years) < 5:
                QMessageBox.warning(self, '警告','请选择至少5个年份以保证预测的准确性。')
                return

            for dic in datas:
                for key in list(dic.keys())[8 : ]:
                    if key not in years:
                        dic.pop(key)

            self.tbw_process_class.clear()
            self.tbw_process_class.add_data(datas, years)

            values = []
            for i in range(len(datas)):
                temp = np.array([])
                for year in years:
                    # if year in datas[i].keys():
                    try:
                        temp = np.append(temp, float(datas[i][year]))
                    except TypeError:
                        temp = np.append(temp, 0)

                   
                values.append(temp)
            values = np.array(values)  # 作为预测函数的y数据列表
            predict_year = self.LineEdit_predict_year.text().split(',')
            
            filename = self.lsw_forecast.selectedItems()[0].text()
            try:
                func = importlib.import_module('funcs.forecast_funcs.%s' % filename).forecast
            except:
                print('加载预测函数失败！')
                return

            try:
                pmax = int(max(predict_year))
                for item in predict_year:
                    int(item)
                # print(pmax)
            except:
                QMessageBox.warning(self, "警告", "请输入合法预测年份")
                return

            # results = func(years, values, predict_year)
            try:
                # pmax = int(max(predict_year))
                results = func(years, values, predict_year)
            except:
                print('预测函数中存在bug！')
                return
            # print(results)
            # print("\n\n")
            # print(pmax)
            # try:
            self.tbw_process_class.display_forecast(results, pmax)
            
            self.tbw_process_class.widget.clearSelection()

    # 预测分析-分析函数
    @QtCore.pyqtSlot()
    def on_btn_process_add_analyze_clicked(self):
        self.load_analyze_func()

    # 预测分析-分析
    @QtCore.pyqtSlot()
    def on_btn_process_analyze_clicked(self):
        # print("here")
        self.selected_year_process = []  # 存放当前在gridLayout中被选中的年份
        self.selected_region_process = []  # 存放当前在gridLayout中被选中的地域
        root = self.trw_process.invisibleRootItem()
        self.selected_category_process = self.get_treeWidget_selected(root)  # 存放当前在tableWidget中被选中的指标
        region_checkboxes = self.get_gridLayout_widgets(self.glo_city_process)
        for checkbox in region_checkboxes:
            if checkbox.isChecked():
                self.selected_region_process.append(checkbox.text())
        year_checkboxes = self.get_gridLayout_widgets(self.glo_year_process)
        for checkbox in year_checkboxes:
            if checkbox.isChecked():
                self.selected_year_process.append(checkbox.text())

        if not self.selected_region_process or not self.selected_year_process or not self.selected_category_process:
            return

        temp_str = '`编码`,`地域`,'
        condition_str = '('
        for region_index in range(len(self.selected_region_process)):
            if region_index != len(self.selected_region_process) - 1:
                condition_str += '`地域`=' + "'%s' or " % self.selected_region_process[region_index]
            else:
                condition_str += '`地域`=' + "'%s') and (" % self.selected_region_process[region_index]
        for category_index in range(len(self.selected_category_process)):
            if category_index != len(self.selected_category_process) - 1:
                condition_str += '`编码`=' + "%d or " % int(self.selected_category_process[category_index])
            else:
                condition_str += '`编码`=' + "%d)" % int(self.selected_category_process[category_index])
        for cat_index in range(1, 7):
            temp_str += '`级别%d`,' % cat_index
       
        temp_str += '`年份`'
        query = 'select %s from `Data` where %s' % (temp_str, condition_str)

        # print(query)
        try:
            datas = self.database.get_data(query)
        except EXECUTE_FAILURE as exef:
            QMessageBox.critical(self, "错误", exef.__str__())
            return
            
        # year = ['2010', '2011', '2012', '2013', '2014', '2015', '2016', '2017', '2018', '2019']
        year = self.year
        for item in year:
            if item not in self.selected_year_process:
                for dic in datas:
                    if item in dic.keys():
                        dic.pop(item)
        # print("************************")
        # print(datas)
        if datas and self.lsw_analyze.selectedItems():
            self.tbw_process_class.clear()
            self.tbw_process_class.add_data(datas, self.selected_year_process)
       

            filename = self.lsw_analyze.selectedItems()[0].text()
            func = importlib.import_module('funcs.analyze_funcs.%s' % filename).analyze
            try:
                func = importlib.import_module('funcs.analyze_funcs.%s' % filename).analyze
            except:
                print('加载分析函数失败！')
                return

            for item in datas:
                item['编码'] = str(item['编码'])
            # print("datas")
            # print(len(datas))
            # print(datas)

            # results, headers = func(datas)
            try:
                results, headers = func(datas)
            except:
                print('分析函数中存在bug！')
                return

            # self.tbw_process_class.display_analyze(results,headers)
            print("call outside try\n\n")
            try:
                print("call inside try")
                self.tbw_process_class.display_analyze(results, headers)
            except:
                print('分析函数返回数据的形式出错！')
                return
            self.tbw_process_class.widget.clearSelection()

    # 预测分析-清空 清空tableWidget
    @QtCore.pyqtSlot()
    def on_btn_process_clear_clicked(self):
        self.tbw_process_class.clear()

    # 预测分析-导出 导出excel
    @QtCore.pyqtSlot()
    def on_btn_process_export_clicked(self):
        if self.tbw_process_class.datas:
            date = time.strftime("%Y_%m_%d_%H-%M-%S", time.localtime())
            # try to fix it (whh)
            if self.tbw_process_class.getMode() == 0:
                filename = QtWidgets.QFileDialog.getSaveFileName(self, "导出Excel", date, "Excel文件 (*.xls)")
            else:
                filename = QtWidgets.QFileDialog.getSaveFileName(self, "导出Excel", date, "Excel文件 (*.xlsx)")
            if filename[0] != '':
                self.tbw_process_class.export_excel(filename[0])

    # ======================================================预测分析

    # ======================================================绘图
    # 绘图-导入数据
    @QtCore.pyqtSlot()
    def on_btn_draw_import_2_clicked(self):
        filename = QtWidgets.QFileDialog.getOpenFileName(self, "打开Excel文件", "", "Excel文件 (*.xls *.xlsx)")
        if filename[0] != '':
            self.draw_excel_path = filename[0]
            self.draw_class.read_data(self.draw_excel_path)

            dialog_2 = Dialog_2()
            dialog_2.display(self.draw_excel_path)
            dialog_2.exec_()

    # 绘图-显示数据
    @QtCore.pyqtSlot()
    def on_btn_draw_show_2_clicked(self):
        if self.draw_excel_path != '' and not self.usedb:
            self.dialog_2 = Dialog_2()
            self.dialog_2.display(self.draw_excel_path)
            self.dialog_2.exec_()

        if self.draw_data and self.usedb:
            self.dialog_2 = Dialog_2()
            self.dialog_2.display_2(self.draw_data)
            self.dialog_2.exec_()
                

    # 绘图-绘图函数
    @QtCore.pyqtSlot()
    def on_btn_draw_add_plot_2_clicked(self):
        self.load_plot_func()

    # 绘图-导出
    @QtCore.pyqtSlot()
    def on_btn_draw_export_2_clicked(self):
        if self.draw_class.pixmap:
            date = time.strftime("%Y_%m_%d_%H-%M-%S", time.localtime())
            filename = QtWidgets.QFileDialog.getSaveFileName(self, "导出图片", date, "图片文件 (*.png *.jpg *.html)")
            if filename[0] != '':
                self.draw_class.save(filename[0])

    # 绘图-绘图
    @QtCore.pyqtSlot()
    def on_btn_draw_plot_2_clicked(self):
        if self.usedb:
            self.query_for_draw()
            self.draw_class.years = self.selected_year_draw
            self.draw_class.datas = self.draw_data
            des = ''
            for i in range(len(self.selected_region_draw)):
                des += self.selected_region_draw[i]
                if i != len(self.selected_region_draw) - 1:
                    des += ','
            des += '数据'

            self.draw_class.description = des
          
        if self.draw_class.datas and self.lsw_plot_2.selectedItems():
            self.usedb = True
           
            filename = self.lsw_plot_2.selectedItems()[0].text()
            try:
                func = importlib.import_module('funcs.plot_funcs.%s' % filename).plot
            except:
                print('加载绘图函数失败！')
                return
           
            font = QtGui.QFont()
            font.setFamily("微软雅黑")
            font.setPointSize(10)
            
            try:
                save_path = self.draw_class.plot_normal(func)
            except:
                print('绘图函数中存在bug！')
                return
    
            messageBox = QMessageBox(QMessageBox.Information, '提示', '能流图已绘制完成并保存在路径\"%s\"中' % save_path)
            messageBox.setFont(font)
            
            # messageBox = QMessageBox(QMessageBox.warning, "提示", "能流图已绘制完成并保存在路径")
            Qyes = messageBox.addButton(self.tr("打开图片"), QMessageBox.YesRole)
            Qno = messageBox.addButton(self.tr("好的"), QMessageBox.NoRole)

            messageBox.exec_()
            if messageBox.clickedButton() == Qyes:
                win32api.ShellExecute(0, 'open', save_path, '', '', 1)
            else:
                return 


            try:
                save_path = self.draw_class.display()
            except:
                print('绘图函数返回数据的形式出错！')
                return

    # ======================================================绘图

    # ======================================================窗口
    # 关闭窗口
    @QtCore.pyqtSlot()
    def on_bnt_close_clicked(self):
        self.close()

    # 最小化窗口
    @QtCore.pyqtSlot()
    def on_bnt_minimize_clicked(self):
        self.showMinimized()

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


    # ================================================年鉴窗口
    
    # 导入模板
    @QtCore.pyqtSlot()
    def on_btn_add_template_clicked(self):
        self.load_templates()


    # 调用现有代码生成结构化的excel
    @QtCore.pyqtSlot()
    def on_btn_generate_excel_clicked(self):    
        # import auto, auto2
        # 数据合法性判定（todo）
        dir = self.lineEdit_dir.text()
        city = self.lineEdit_city.text()
        self.citypath = city
        year = self.lineEdit_year.text()
        threshold = self.lineEdit_threshold.text()
        select_item = self.template_list.selectedItems()

        lose = []
        if(not dir):
            lose.append('年鉴目录')
        if(not city):
            lose.append('输出路径')
        if(not year):
            lose.append('年鉴年份')
        if(not threshold):
            lose.append('匹配参数')
        if(not select_item):
            lose.append('模板')

        if(lose):
            message = ''
            for i in range(len(lose)):
                message += "\"%s\"" % lose[i]
                if i != len(lose) - 1:
                    message += '、'
            message += '不能为空'
            QMessageBox.warning(self, '警告', message)

        else:    
            self.mythread = MyThread(dir, city, year, threshold, select_item)
            self.mythread.signal.connect(self.callback)
            self.mythread.finish.connect(self.close_ui)
            self.mythread.start()
            self.calculating_ui = calculating()
            self.calculating_ui.exec_()
        

    def close_ui(self):
        self.calculating_ui.finish_cal(self.citypath)

    def callback(self, i):
        self.calculating_ui.setpercentage(i)


    @QtCore.pyqtSlot()
    def on_btn_generate_excel2_clicked(self):   
        # import auto, auto2
        # 数据合法性判定（todo）
        dir = self.lineEdit_dir.text()
        city = self.lineEdit_city.text()
        year = self.lineEdit_year.text()
        threshold = self.lineEdit_threshold.text()
        select_item = self.template_list.selectedItems()
        

    def browse_path(self):
        filepath = QtWidgets.QFileDialog.getExistingDirectory(self, "选择年鉴目录")
        self.lineEdit_dir.setText(filepath)
        # print(filepath)

    def browse_output(self):
        filepath = QtWidgets.QFileDialog.getExistingDirectory(self, "选择输出路径")
        self.lineEdit_city.setText(filepath)
        # print(filepath)


class MyThread(QThread):
    from auto import auto
    import auto2
    signal = pyqtSignal(float)
    finish = pyqtSignal()
    def __init__(self, dir, city, year, threshold, select_item):
        super(MyThread, self).__init__()
        self.dir = dir
        self.city = city
        self.year = year
        self.threshold = threshold
        self.select_item = select_item
        self.thisauto = auto()
        self.thisauto.signal.connect(self.callback)

    def run(self):
        import auto2
        if self.dir != "" and self.city != "" and self.year != "" and self.threshold != "":
            if self.select_item:
                template_name = self.select_item[0].text()
                # self.thisauto = auto()
                self.thisauto.generate_without_mapping(self.dir, self.city, self.year, 'templates\\' + template_name+'.xls', template_name, float(self.threshold))
                # self.thisauto.signal.connect(self.callback)
                self.finish.emit()

    def callback(self, i):
        self.signal.emit(i)

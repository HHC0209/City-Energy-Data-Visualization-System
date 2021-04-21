import modify
import getData
import shutil
import xlwt
import xlrd
import os
import collect
from match import match
from PyQt5.QtCore import *
# dir = r'D:\project\energy data\城市\天津\天津统计年鉴2019（EXCEL版）'  # 某年鉴的表格文件夹
# city = './天津'  # 填入你在做的城市名称
# year = '2019'  # 年鉴年份
# baseexcel = '统一目录.xlsx'  # 统一目录
class auto(QObject):
    signal = pyqtSignal(float)

    def __init__(self):
        super(QObject, self).__init__()
        self.thismatch = match()
        self.thismatch.signal.connect(self.callback)

    def generate_without_mapping(self, dir, city, year, baseexcel, city_name, match_threshold):
        # print("auto1")
        if not os.path.exists(city):
            os.mkdir(city)
        cityAndYear = city + "\\" + city_name
        if not os.path.exists(cityAndYear):
            os.mkdir(cityAndYear)
        excelDir = cityAndYear + "\\" + 'excel'  # 存放此年鉴所有表格的文件夹
        collectDir = cityAndYear + "\\" + 'collect'  # 存放匹配到表名的表格的文件夹，即符合要求的表格，其中name.txt为存放各表格表名的文本文档，方便查看
        modifyDir = cityAndYear + "\\" + 'modify'  # 存放调整格式后的表格的文件夹
        dataExcel = cityAndYear + "\\" + 'data.xls'  # 存放爬取下来的数据的表格
        matchExcel = cityAndYear + "\\" + city_name + '.xls'  # 存放匹配到的数据

        collect.collectExcel(dir, excelDir, baseexcel, collectDir)
        modify.modifyDir(collectDir, modifyDir)
        getData.getDataDir(modifyDir, dataExcel)

        # self.thismatch = match()
        self.thismatch.match(baseexcel, dataExcel, matchExcel, year, city_name, threshold=match_threshold)  # 这一步运行非常耗时间，可考虑单独拿出来独立运行
        

    def callback(self, i):
        # print("收到")
        self.signal.emit(i)
import modify
import getData
import shutil
import xlwt
import xlrd
import os
import collect
import match

# dir = r'D:\project\energy data\城市\苏州市和无锡市的统计年鉴\无锡统计年鉴2018(光盘版)\zk\html'  # 某年鉴的表格文件夹
# city = './无锡'  # 填入你在做的城市名称
# year = '2018'  # 年鉴年份
# baseexcel = r'match.xls'  # 调整之后的match.xls文件路径

def generate_with_mapping(dir, city, year, baseexcel, city_name,match_threshold):
    print("auto2")
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
    

    collect.collectExcel(dir, excelDir, baseexcel, collectDir, pos=7, threshold=0.49, num=6)
    modify.modifyDir(collectDir, modifyDir)
    getData.getDataDir(modifyDir, dataExcel)

    match.match2(baseexcel, dataExcel, matchExcel, year, city_name, threshold=match_threshold)  # 这一步运行非常耗时间，可考虑单独拿出来独立运行
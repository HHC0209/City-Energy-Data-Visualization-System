from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import QSize
from PyQt5.QtCore import Qt
import plotly.graph_objects as go
import plotly.offline as pyof
import xlrd
import shutil
from showGraph import showGraph
from PyQt5.QtCore import *
# import seaborn as sns


class Draw(QObject):
    signal = pyqtSignal()
    def __init__(self, lab_draw):
        super(QObject, self).__init__()
        self.flag = -1
        self.widget = lab_draw
        self.datas = []
        self.years = []
        self.description = ''
        self.pixmap = None

    def read_data(self, draw_excel_path):
        self.signal.emit()
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
            self.read_data_0(worksheet)
            self.description = '%s数据' % region_str
        else:
            self.read_data_1(worksheet)
            try:
                year = int(header[1])
                region = header[0]
                self.description = '%d%s 能流图数据' % (year, region)
            except:
                self.description = "能流图"

    def read_data_0(self, worksheet):
        header = worksheet.row_values(0)  # 读取excel第1行的header
        self.flag = 0
        # 数据从excel的第2行开始
        self.datas = []
        for row_num in range(1, worksheet.nrows):
            row = worksheet.row_values(row_num)
            temp = {}
            for col_num in range(len(row)):
                if col_num == 0:
                    temp['编码'] = str(int(row[col_num]))
                elif col_num >= 8:
                    try:
                        temp[header[col_num]] = round(float(row[col_num]), 3)
                    except:
                        temp[header[col_num]] = 0
                else:
                    temp[header[col_num]] = row[col_num]
            self.datas.append(temp)
        self.years = [year for year in header[8:]]

    def plot_normal(self, func):
        self.widget.clear()
        year_list = []
        for item in self.years:
            if item in self.datas[0].keys():
                year_list.append(item)
        
        # print(self.years)
        # print(year_list)
        # print(self.datas)
        # print(len(self.data))

        return func(self.datas, year_list, self.description)

    def display(self):
        self.pixmap = QPixmap('graphs\\%s.png' % self.description)
        # 自适应大小
        width = self.pixmap.width()
        height = self.pixmap.height()
        # scaleh = 760 / height
        # scalew = 1200 / width
        # scale = 630 / height
        scale = 861 / width
        size = QSize(width * scale, height * scale)
        # size = QSize(width*scalew, height*scaleh)
        self.pixmap = self.pixmap.scaled(size, Qt.KeepAspectRatio)
        self.widget.setPixmap(self.pixmap)
        # self.widget.setScaledContents(True)
        # graph_dlg = showGraph(self.widget)
        # graph_dlg.display()
        # graph_dlg.exec_()

    def read_data_1(self, worksheet):
        header = worksheet.row_values(0)  # 读取excel第1行的header
        self.flag = 0
        self.years = None
        # 数据从excel的第2行开始
        self.datas = []
        # if worksheet.ncols == 3:
        #     for col_num in range(worksheet.ncols):
        #         tmp = []
        #         col = worksheet.col_values(col_num)
        #         for index in range(1, worksheet.nrows):
        #             tmp.append(col[index])
        #         self.datas.append(tmp)

        # c1, c2, c3 = [], [], []
        # flag_of_table = [[1,1,1,1,0,0,1,1],[0,0,0,0,1,0,0,0],[0,0,0,0,1,0,0,0],[0,0,0,0,1,0,0,0],[0,0,0,0,1,0,0,0],[0,0,0,0,0,0,1,0],[0,0,0,0,1,1,0,0],[0,0,0,0,0,1,0,0],[0,0,0,0,1,0,0,0],[0,0,0,0,0,1,0,0],[1,1,1,0,0,1,0,1],[0,0,0,0,1,0,0,0],[0,0,0,0,1,0,0,0],[1,1,1,0,0,0,0,1],[0,0,0,0,0,1,0,0],[0,0,0,0,0,1,0,0],[1,0,0,0,0,0,0,0],[0,0,0,1,0,0,0,1],[0,0,0,1,0,0,0,0],[1,1,0,0,1,0,0,0],[1,1,1,1,1,1,1,1],[1,1,1,1,1,1,0,0],[1,1,0,0,1,0,0,0],[1,1,1,0,1,1,0,0],[1,1,1,0,1,1,0,0]]
        # for i in range(len(flag_of_table)):
        #     row = worksheet.row_values(i + 1)
        #     for j in range(len(flag_of_table[0])):
        #         if flag_of_table[i][j] == 1:
        #             sign = row[0][-2]
        #             if sign == "+":
        #                 c1.append(row[0][:-3])
        #                 c2.append(header[j+1])
        #             else:
        #                 c1.append(row[0][:-3]) 
        #                 c2.append(header[j+1])
        #             c3.append(row[j+1]) 
        # self.datas = [c1, c2, c3]

        c1 = ['煤炭', '煤炭', '煤炭', '煤炭', '煤炭', '煤炭', '煤炭', '煤炭', '煤炭', '石油', '石油', '石油', '石油', '石油', '石油', '石油', '天然气', '天然气', '天然气', '天然气', '天然气', '天然气', '炼焦过程', '炼焦过程', '焦炭', '焦炭', '水力发电', '核能发电', '太阳能发电', '风力发电', '垃圾和生物质用于燃料', '焚烧垃圾', '焚烧垃圾', '焚烧垃圾', '外来电', '电能生产', '电能生产', '电能', '电能', '电能', '电能', '电能', '电能', '热能', '热能生产', '热能生产', '热能', '热能', '热能', '热能', '炼焦过程', '其它焦化产品', '其它焦化产品', '其它焦化产品']
        c2 = ['电能生产', '热能生产', '炼焦过程', '农业消费', '规上工业消费', '规下工业消费', '建筑业消费', '第三产业消费', '居民生活消费', '电能生产', '农业消费', '规上工业消费', '规下工业消费', '建筑业消费', '第三产业消费', '居民生活消费', '电能生产', '热能生产', '规上工业消费', '规下工业消费', '第三产业消费', '居民生活消费', '焦炭', '转换损失', '规上工业消费', '规下工业消费', '电能', '电能', '电能', '电能', '焚烧垃圾', '电能', '热能', '转换损失', '电能', '电能', '转换损失', '农业消费', '规上工业消费', '规下工业消费', '建筑业消费', '第三产业消费', '居民生活消费', '电能生产', '热能', '转换损失', '规上工业消费', '规下工业消费', '第三产业消费', '居民生活消费', '其它焦化产品', '规上工业消费', '电能生产', '热能生产']
        c3 = []
        pairs = [[10, 0], [13, 0], [16, 0], [19, 0], [20, 0], [21, 0], [22, 0], [23, 0], [24, 0], [10, 1], [19, 1], [20, 1], [21, 1], [22, 1], [23, 1], [24, 1], [10, 2], [13, 2], [20, 2], [21, 2], [23, 2], [24, 2], [17, 3], [18, 3], [20, 3], [21, 3], [1, 4], [2, 4], [3, 4], [4, 4], [5, 6], [6, 4], [6, 5], [7, 5], [8, 4], [11, 4], [12, 4], [19, 4], [20, 4], [21, 4], [22, 4], [23, 4], [24, 4], [10, 5], [14, 5], [15, 5], [20, 5], [21, 5], [23, 5], [24, 5], [17, 7], [20, 7], [10, 7], [13, 7]]

        # print(worksheet.row_values(2)[0])
        for pair in pairs:
            c3.append(worksheet.row_values(pair[0] + 1)[pair[1] + 1])
        print(len(c1), len(c2), len(c3))
        self.datas = [c1, c2, c3]



    def plot_energy_flow(self):
        # todo
        pass

    def save(self, filepath):
        try:
            source = 'graphs\\%s.png' % self.description
            shutil.copy(source, filepath)
        except:
            source = 'graphs\\test.html'
            shutil.copy(source, filepath)


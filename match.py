import os
import pandas as pd
import jieba
import numpy as np
import gensim
from scipy.linalg import norm
import re
import difflib
import xlwt
import xlrd
from tqdm import tqdm
from PyQt5.QtCore import *
class match(QObject):

    signal = pyqtSignal(float)
    def __init__(self):
        super(QObject, self).__init__()

# 调用NLP模型的语义相似度算法
    def vector_similarity(self, s1, s2):
        # 推荐0.85-0.90为好
        def sentence_vector(s):
            words = jieba.lcut(s)
            v = np.zeros(64)
            for word in words:
                v += model[word]
            v /= len(words)
            return v

        v1, v2 = sentence_vector(s1), sentence_vector(s2)
        # print(v1, v2)
        return np.dot(v1, v2) / (norm(v1) * norm(v2))


    # 字符串比较语义相似度算法
    def string_similar(self, s1, s2):
        # 推荐0.75-0.8为好
        return difflib.SequenceMatcher(None, s1, s2).quick_ratio()


    # 余弦距离语义相似度算法
    def cos_similar(self, s1, s2):
        # 推荐0.70-0.75为好
        list1 = list(jieba.cut(s1))
        list2 = list(jieba.cut(s2))
        key_word = list(set(list1 + list2))
        word_vector1 = np.zeros(len(key_word))
        word_vector2 = np.zeros(len(key_word))
        for i in range(len(key_word)):
            # 遍历key_word中每个词在句子中的出现次数
            for j in range(len(list1)):
                if key_word[i] == list1[j]:
                    word_vector1[i] += 1
            for k in range(len(list2)):
                if key_word[i] == list2[k]:
                    word_vector2[i] += 1
        dist = float(np.dot(word_vector1, word_vector2) / (np.linalg.norm(word_vector1) * np.linalg.norm(word_vector2)))
        return dist


    def regularization(self, str1):
        word = str1
        # r = '[a-zA-Z’!"#$%&\'()*+./<=>?@。#?★…【】《》？“”‘’！[\\]^_`{|}~（）、]+'  # 仅去除各种符号，不能去除数字，把-去除掉了
        r = '[a-zA-Z]+'
        word = re.sub(r, '', word).replace('\n', '').replace(' ', '').replace('\u3000', '')
        if word == '':
            word = '空'
        return word

    # baseexcel:统一目录.xlsx data:data.xls  matchExcel:存放匹配结果的表格 year:年鉴年份
    def match(self, baseexcel, data, matchExcel, year, city_name, threshold=0.535):
        # signal = pyqtSignal(i)
        rbBase = xlrd.open_workbook(baseexcel)
        tableBase = rbBase.sheets()[0]
        rowBase = tableBase.nrows
        colBase = tableBase.ncols

        rbData = xlrd.open_workbook(data)
        tableData = rbData.sheets()[0]
        rowData = tableData.nrows
        colData = tableData.ncols

        workbook = xlwt.Workbook(encoding='utf-8')
        worksheet = workbook.add_sheet('Sheet1')

        nomatchExcel = "no" + matchExcel
        workbook1 = xlwt.Workbook(encoding='utf-8')
        worksheet1 = workbook1.add_sheet('Sheet1')

        # 添加heading
        worksheet.write(0,0,"编码")
        worksheet.write(0,1,"地域")
        worksheet.write(0,2,"级别1")
        worksheet.write(0,3,"级别2")
        worksheet.write(0,4,"级别3")
        worksheet.write(0,5,"级别4")
        worksheet.write(0,6,"级别5")
        worksheet.write(0,7,"级别6")
        worksheet.write(0,8,str(year))
        worksheet.write(0,9,"相似度")
        cnt_row = 1
        for i in (range(1, rowBase)):  # 第一列是关键字不需要匹配
            # print(i)
            # print(i / (rowBase - 1))
            self.signal.emit(i / (rowBase - 1))
            s1 = ""
            s1 = s1 + str(float(year) - 1)  # 加入年份进行匹配，不知道会不会影响准确率
            for c in range(2, 7):  # 第2到7列
                if tableBase.cell_value(i, c) != tableBase.cell_value(i, c - 1):
                    s1 = s1 + str(tableBase.cell_value(i, c))
            s1 = self.regularization(s1)
            scorelis = []
            for j in range(rowData):
                s2 = ""
                for c in range(3):  # 第1到3列
                    s2 = s2 + str(tableData.cell_value(j, c))
                s2 = self.regularization(s2)
                if "三资" in s2:  # 无锡2019年的三资表格影响了匹配准确率
                    s2 = "三资"
                score = self.string_similar(s1, s2)
                scorelis.append(score)
            scoremax = max(scorelis)
            indmax = scorelis.index(scoremax)
            # 修改了输出文件的格式（whh）
            # if scoremax >= threshold:
            #     for k in range(colBase):
            #         worksheet.write(i, k, tableBase.cell_value(i, k))
            #     for k in range(colBase, colBase + colData):
            #         worksheet.write(i, k, tableData.cell_value(indmax, k - colBase))
            #     worksheet.write(i, colBase + colData, scoremax)
            #     # count = count + 1
            # else:
            #     for k in range(colBase):
            #         worksheet.write(i, k, tableBase.cell_value(i, k))
            #     # count = count + 1

            #     for k in range(colBase):
            #         worksheet1.write(i, k, tableBase.cell_value(i, k))
            #     for k in range(colBase, colBase + colData):
            #         worksheet1.write(i, k, tableData.cell_value(indmax, k - colBase))
            #     worksheet1.write(i, colBase + colData, scoremax)
            if scoremax >= threshold:
                if tableData.cell_value(indmax, colData-1) != "":
                    worksheet.write(cnt_row, 0, tableBase.cell_value(i, 0))
                    worksheet.write(cnt_row, 1, str(city_name))
                    for k in range(1, colBase):
                        worksheet.write(cnt_row, k+1, str(tableBase.cell_value(i, k)))
                    worksheet.write(cnt_row, colBase+1, tableData.cell_value(indmax, colData-1))
                    worksheet.write(cnt_row, colBase+2, scoremax)
                    cnt_row += 1   

        worksheet.col(0).width = 2000
        worksheet.col(1).width = 3000
        worksheet.col(2).width = 6000
        worksheet.col(3).width = 6000
        worksheet.col(4).width = 6000
        worksheet.col(5).width = 6000
        worksheet.col(6).width = 6000
        worksheet.col(7).width = 8000  # Data
        worksheet.col(8).width = 8000
        worksheet.col(9).width = 8000
        worksheet.col(10).width = 4000
        worksheet.col(11).width = 4000
        workbook.save(matchExcel)


    def match2(self, baseexcel, data, matchExcel, year, city_name, threshold=0.535):
        rbBase = xlrd.open_workbook(baseexcel)
        tableBase = rbBase.sheets()[0]
        rowBase = tableBase.nrows
        # colBase = tableBase.ncols
        colBase = 7

        rbData = xlrd.open_workbook(data)
        tableData = rbData.sheets()[0]
        rowData = tableData.nrows
        colData = tableData.ncols

        workbook = xlwt.Workbook(encoding='utf-8')
        worksheet = workbook.add_sheet('Sheet1')

        nomatchExcel = "no" + matchExcel
        workbook1 = xlwt.Workbook(encoding='utf-8')
        worksheet1 = workbook1.add_sheet('Sheet1')


        # 添加heading
        worksheet.write(0,0,"编码")
        worksheet.write(0,1,"地域")
        worksheet.write(0,2,"级别1")
        worksheet.write(0,3,"级别2")
        worksheet.write(0,4,"级别3")
        worksheet.write(0,5,"级别4")
        worksheet.write(0,6,"级别5")
        worksheet.write(0,7,"级别6")
        worksheet.write(0,8,str(year))
        worksheet.write(0,9,"相似度")
        cnt_row = 1
        for i in (range(1, rowBase)):  # 第一列是关键字不需要匹配
            print(i)
            s1 = ""
            s1 = s1 + str(float(year) - 1)  # 加入年份进行匹配，不知道会不会影响准确率
            for c in range(7, 10):  # 第7到10列
                if tableBase.cell_value(i, c) != tableBase.cell_value(i, c - 1):
                    s1 = s1 + str(tableBase.cell_value(i, c))
            s1 = self.regularization(s1)
            scorelis = []
            for j in range(rowData):
                s2 = ""
                for c in range(3):  # 第1到3列
                    s2 = s2 + str(tableData.cell_value(j, c))
                s2 = self.regularization(s2)
                if "三资" in s2:  # 无锡2019年的三资表格影响了匹配准确率
                    s2 = "三资"
                score = self.string_similar(s1, s2)
                scorelis.append(score)
            scoremax = max(scorelis)
            indmax = scorelis.index(scoremax)
            # if scoremax >= threshold:
            #     for k in range(colBase):
            #         worksheet.write(i, k, tableBase.cell_value(i, k))
            #     for k in range(colBase, colBase + colData):
            #         worksheet.write(i, k, tableData.cell_value(indmax, k - colBase))
            #     worksheet.write(i, colBase + colData, scoremax)
            #     # count = count + 1
            # else:
            #     for k in range(colBase):
            #         worksheet.write(i, k, tableBase.cell_value(i, k))
            #     # count = count + 1

            #     for k in range(colBase):
            #         worksheet1.write(i, k, tableBase.cell_value(i, k))
            #     for k in range(colBase, colBase + colData):
            #         worksheet1.write(i, k, tableData.cell_value(indmax, k - colBase))
            #     worksheet1.write(i, colBase + colData, scoremax)
            if scoremax >= threshold:
                if tableData.cell_value(indmax, colData-1) != "":
                    worksheet.write(cnt_row, 0, tableBase.cell_value(i, 0))
                    worksheet.write(cnt_row, 1, str(city_name))
                    for k in range(1, colBase):
                        worksheet.write(cnt_row, k+1, str(tableBase.cell_value(i, k)))
                    worksheet.write(cnt_row, colBase+1, tableData.cell_value(indmax, colData-1))
                    worksheet.write(cnt_row, colBase+2, scoremax)
                    cnt_row += 1

        worksheet.col(0).width = 2000
        worksheet.col(1).width = 3000
        worksheet.col(2).width = 6000
        worksheet.col(3).width = 6000
        worksheet.col(4).width = 6000
        worksheet.col(5).width = 6000
        worksheet.col(6).width = 6000
        worksheet.col(7).width = 8000  # Data
        worksheet.col(8).width = 8000
        worksheet.col(9).width = 8000
        worksheet.col(10).width = 4000
        worksheet.col(11).width = 4000
        workbook.save(matchExcel)

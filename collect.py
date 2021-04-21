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
import shutil
from xlutils.copy import copy


# 打开一个excel文件，返回这个文件的文件名
def nameofexcel(filepath):
    # try:
    #     df = pd.read_excel(filepath)
    # except:
    #     return "error"
    # else:
    #     pass
    df = pd.read_excel(filepath)
    col1 = df.columns.values[0]  # 提取一级标题
    # print(col1)
    # 去掉空格
    # try:
    #     rb = xlrd.open_workbook(filepath)
    # except:
    #     wb = copy(rb)
    #     ws = wb.get_sheet(0)
    #     wb.save("temp.xls")
    #     rb = xlrd.open_workbook("temp.xls")
    # else:
    #     pass
    # table = rb.sheets()[0]
    # col1 = table.cell_value(0, 0)
    try:
        ind = col1.index(' ')
    except:
        pass
    else:
        col1 = col1[ind:]

    # 去掉开头末尾的空格
    col1 = col1.lstrip()
    col1 = col1.rstrip()
    # 去掉（和(  # 有些情况是在中间有括号  # 考虑）是在末尾的情况
    if not col1 == "":
        if col1[-1] == ')' or col1[-1] == '）':
            try:
                ind = col1.index('（')
            except:
                try:
                    ind = col1.index('(')
                except:
                    pass
                else:
                    col1 = col1[:ind]
            else:
                col1 = col1[:ind]
    return col1


# 调用NLP模型的语义相似度算法
def vector_similarity(s1, s2):
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
def string_similar(s1, s2):
    # 推荐0.75-0.8为好
    return difflib.SequenceMatcher(None, s1, s2).quick_ratio()


# 余弦距离语义相似度算法
def cos_similar(s1, s2):
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


# 获得某文件夹下所有的excel文件的路径
def getAllExcel(dir):
    lis1 = []

    def getExcel(dir):
        templis = os.listdir(dir)
        for i in templis:
            if not os.path.isdir(dir + "\\" + i):
                if i[-3:] == 'xls' or i[-4:] == 'xlsx':
                    lis1.append(dir + "\\" + i)
            else:
                getExcel(dir + "\\" + i)

    getExcel(dir)
    return lis1


# 将包含文件路径的链表中的所有excel文件都复制到另一个文件夹下
def copyFile(lis, targetdir):
    if not os.path.exists(targetdir):
        os.mkdir(targetdir)
    for i in lis:
        shutil.copy(i, targetdir)


# 给定一个链表，找出最大的N个值，判断是否超过阈值，返回它们的下标
def findMaxIndex(lis, threshold=0.0, num=1):
    templis = []
    for i in range(len(lis)):
        templis.append(lis[i])
    index = []
    index0 = []  # 删去阈值的
    for i in range(num):
        temp = max(templis)
        try:
            tempindex = lis.index(temp)
        except:
            continue
        else:
            templis[tempindex] = 0  # 如果出现重复的几个都可以记录下来
            index.append(tempindex)
    for i in range(len(index)):
        # print(i, index[i], lis[index[i]], threshold)
        if(lis[index[i]] >= threshold):
            index0.append(index[i])
    # templis.sort(reverse=True)
    return index0


# 给定一个文件夹，里面都是excel文件，返回这些excel文件的路径以及它们的表名链表
def getPathAndName(dir):
    xlslist = []
    xlslistname = []
    filelist = os.listdir(dir)
    for i in filelist:
        if i[-3:] == 'xls' or i[-4:] == 'xlsx':
            xlslist.append(dir + "\\" + i)
            xlslistname.append(nameofexcel(dir + "\\" + i))
    return xlslist, xlslistname


# 返回统一目录的各个表格的名称，默认其出现在第3列
def getBaseName(baseexcel, pos=2):
    dataBase = pd.read_excel(baseexcel)
    columns = len(dataBase.iloc[:, pos].values)
    rows = len(dataBase.iloc[pos, :].values)
    excelNameBase = dataBase.iloc[:, pos].values
    return excelNameBase


# 输入统一目录表名和年鉴表名链表，返回正则化去重复的统一目录文件表名和正则化表名链表
def regularization(excelNameBase, xlslistname):
    # r = '[a-zA-Z’!"#$%&\'()*+-./<=>?@。#?★…【】《》？“”‘’！[\\]^_`{|}~（）、0-9]+'
    r = '[a-zA-Z]+'
    excelNameBase0 = []  # 正则化后的统一目录里的表名
    excelNameBase00 = []  # 去重复后的正则化后的统一目录表名
    xlslistname0 = []  # 正则化后的年鉴的表名
    exx = set()
    for word in excelNameBase:  # 可以在这一步实现去重复，目前不选择这个方案
        word = str(word)
        word = re.sub(r, '', word).replace('\n', '').replace(' ', '').replace('\u3000', '')
        if word == '':
            word = '空'
        excelNameBase0.append(word)
    for word in xlslistname:
        word = re.sub(r, '', word).replace('\n', '').replace(' ', '').replace('\u3000', '')
        if word == '':
            word = '空'
        xlslistname0.append(word)
    excelNameBase00.append(excelNameBase0[0])
    # 去重复选择用set会更好，但是set的顺序会出现变化
    for i in range(1, len(excelNameBase0)):
        if excelNameBase0[i] != excelNameBase0[i - 1]:
            excelNameBase00.append(excelNameBase0[i])
    return excelNameBase00, xlslistname0


# 输出各个收集到的表格的表名到一个txt文件中
def printName(lis, file):
    with open(file, "w", encoding='utf-8') as f:
        for i in lis:
            rb = xlrd.open_workbook(i)
            table = rb.sheets()[0]
            f.write(table.cell_value(0, 0) + '\n')


# 处理续表，表名中包含续表关键字的话它的名字会等于上一个表格的名字
def continued(dir):
    xlslist, xlslistname = getPathAndName(dir)
    for i in range(len(xlslistname)):
        if "续表" in xlslistname[i]:
            print("续表")
            rb0 = xlrd.open_workbook(xlslist[i-1])
            table0 = rb0.sheets()[0]
            name = table0.cell_value(0, 0)
            rb = xlrd.open_workbook(xlslist[i])
            table = rb.sheets()[0]
            wb = copy(rb)
            ws = wb.get_sheet(0)
            ws.write(0, 0, name)
            wb.save(xlslist[i])


# 计算相似度，找到符合的表格放到targetDir文件夹下
# dir年鉴表格 excelDir存放年鉴表格的文件夹 baseexcel统一目录表格 targetDir存放符合要求表格的文件夹
def collectExcel(dir, excelDir, baseexcel, targetDir, mode="str", pos=2, threshold=0.39, num=7):
    lis1 = getAllExcel(dir)
    copyFile(lis1, excelDir)
    continued(excelDir)  # 处理续表
    xlslist, xlslistname = getPathAndName(excelDir)
    excelNameBase = getBaseName(baseexcel, pos)
    excelNameBase00, xlslistname0 = regularization(excelNameBase, xlslistname)

    scorelis = []
    if mode == "str":
        for i in range(len(excelNameBase00)):
            templis = []
            for j in range(len(xlslistname0)):
                s1 = excelNameBase00[i]
                s2 = xlslistname0[j]
                score = string_similar(s1, s2)
                templis.append(score)
            scorelis.append(templis)
    elif mode == "cos":
        for i in range(len(excelNameBase00)):
            templis = []
            for j in range(len(xlslistname0)):
                s1 = excelNameBase00[i]
                s2 = xlslistname0[j]
                score = cos_similar(s1, s2)
                templis.append(score)
            scorelis.append(templis)
    elif mode == "model":
        model_file = r'D:\project\energy data\全自动\news_12g_baidubaike_20g_novel_90g_embedding_64.bin'
        model = gensim.models.KeyedVectors.load_word2vec_format(model_file, binary=True)
        for i in range(len(excelNameBase00)):
            templis = []
            for j in range(len(xlslistname0)):
                s1 = excelNameBase00[i]
                s2 = xlslistname0[j]
                score = vector_similarity(s1, s2)
                templis.append(score)
            scorelis.append(templis)
    else:
        pass
    indlis = []
    for i in range(len(scorelis)):
        indlis = indlis + findMaxIndex(scorelis[i], threshold=threshold, num=num)
    filelis = []
    for i in range(len(indlis)):
        filelis.append(xlslist[indlis[i]])
        print(xlslistname[indlis[i]])
    copyFile(filelis, targetDir)
    printName(filelis, targetDir + "\\name.txt")


# s1 = "2018.0规模以上工业主要产品产量布全市"
# s2 = "规模以上工业主要产品产量2018布(万米)(10000)"
# s3 = "规模以上工业主要产品产量1980彩色电视机"
# s4 = "规模以上工业主要产品产量2018彩色电视机"
#
# m1 = string_similar(s1, s3)
# m2 = string_similar(s1, s4)
# print(m1, m2)

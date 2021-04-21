import os
import xlrd
import xlwt
import pandas as pd


def nameofexcel(filepath):
    try:
        df = pd.read_excel(filepath)
    except:
        return "error"
    else:
        pass

    col1 = df.columns.values[0]  # 提取一级标题
    # print(col1)
    # 去掉空格
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


def findFirstNum(table):
    print(table.cell_value(0, 0))
    columns = table.ncols
    rows = table.nrows
    for i in range(rows):
        for j in range(columns):
            if (type(table.cell_value(i, j)) == (float) or type(table.cell_value(i, j)) == (int)):  # 判断是否有数字
                # 判断是不是年份  # 它将年份读为float #可能会漏掉一些数据
                if table.cell_value(i, j) >= 1949 and table.cell_value(i, j) <= 2030:
                    continue
                else:
                    return i, j


def findFirstNumContent(file):
    rb = xlrd.open_workbook(file)
    table = rb.sheets()[0]
    columns = table.ncols
    rows = table.nrows
    firstDataR, firstDataC = findFirstNum(table)
    firstContentR = 0
    firstContentC = firstDataC
    for i in range(rows):
        if (table.cell_value(i, firstDataC) != ""):
            firstContentR = i
            break
    return columns, rows, firstContentR, firstContentC, firstDataR, firstDataC


def delNullData(targetfile, ttargetfile):
    rb = xlrd.open_workbook(targetfile)
    table = rb.sheets()[0]
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('Sheet1')
    count = 0
    for i in range(table.nrows):
        if (table.cell_value(i, 3) != ""):
            worksheet.write(count, 0, table.cell_value(i, 0))
            worksheet.write(count, 1, table.cell_value(i, 1))
            worksheet.write(count, 2, table.cell_value(i, 2))
            worksheet.write(count, 3, table.cell_value(i, 3))
            count = count + 1
    worksheet.col(0).width = 10000
    worksheet.col(1).width = 10000
    worksheet.col(2).width = 15000
    worksheet.col(3).width = 4000
    workbook.save(ttargetfile)


# dir为调整单元格后的文件夹
# tergetfile为爬取所有的数据的总的excel文件
def getDataDir(dir, targetfile):
    filelist = os.listdir(dir)
    xlslist = []
    for i in filelist:
        if i[-3:] == 'xls' or i[-4:] == 'xlsx':
            xlslist.append(dir + "\\" + i)

    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('Sheet1')
    count = 0

    for k in range(len(xlslist)):
        name = nameofexcel(xlslist[k])
        columns, rows, firstContentR, firstContentC, firstDataR, firstDataC = findFirstNumContent(xlslist[k])
        rb = xlrd.open_workbook(xlslist[k])
        table = rb.sheets()[0]
        for i in range(firstDataR, rows):
            for j in range(firstDataC, columns):
                worksheet.write(count, 0, name)
                worksheet.write(count, 1, table.cell_value(i, 0))
                worksheet.write(count, 2, table.cell_value(firstContentR, j))
                worksheet.write(count, 3, table.cell_value(i, j))
                count = count + 1
    # 设置单元格宽度
    worksheet.col(0).width = 10000
    worksheet.col(1).width = 10000
    worksheet.col(2).width = 15000
    worksheet.col(3).width = 4000
    workbook.save(targetfile)
    # delNullData(targetfile, targetfilewithoutnull)

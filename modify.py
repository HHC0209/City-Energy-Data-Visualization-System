import os
import re
import xlrd
from xlutils.copy import copy


# 找到一个表格中第一个数据项的位置，返回它的行，列
def findFirstNum(table):
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


# 对一个excel文件进行处理，合并其目录项的单元格，得到的文件保存在targetfile中
def modify(file, targetfile):
    rb = xlrd.open_workbook(file)
    table = rb.sheets()[0]
    print(table.cell_value(0, 0))
    columns = table.ncols
    rows = table.nrows
    wb = copy(rb)
    ws = wb.get_sheet(0)

    # 查找出现第一个数据的位置
    firstDataR, firstDataC = findFirstNum(table)  # 找出第一个数据所在的位置

    # 在这一列查找找到第一个目录项的起始行位置
    firstContentR = 0
    firstContentC = firstDataC
    for i in range(rows):
        if (table.cell_value(i, firstDataC) != ""):
            firstContentR = i
            break
    if (firstContentC + 1 < columns) and (firstContentR >= 1):  # 第一个数据列的目录项有可能第一行是空单元格，如无锡2019，12-05.xls
        if table.cell_value(firstContentR-1, firstContentC + 1) != "":  # 看旁边那一列的目录起始行是否一样
            firstContentR = firstContentR - 1
    # print(firstDataR, firstDataC, firstContentR, firstContentC)
    # 转换为二维链表的形式方便操作，仅转换数据项上面所有的行
    temp = []
    for i in range(firstDataR):
        tt = []
        for j in range(columns):
            tt.append(table.cell_value(i, j))
        temp.append(tt)
    # 去除英文名
    for i in range(1, columns):
        temp[0][i] = ""
    # # 把所有空的单元格给填补上内容
    # for i in range(firstContentR, firstDataR):
    #     for j in range(firstContentC, columns):
    #         if (temp[i][j] != ""):  # 往右边和下边都填满
    #             for ii in range(i + 1, firstDataR):  # 当前列
    #                 if (temp[ii][j] == ""):
    #                     temp[ii][j] = temp[i][j]
    #                 else:
    #                     break
    #             for jj in range(j + 1, columns):  # 当前行
    #                 if (temp[i][jj] == ""):
    #                     temp[i][jj] = temp[i][j]
    #                 else:
    #                     break
    # 第一行的目录是下面所有行数的和
    # r = '[a-zA-Z’!"#$&\'*+-<=>?@。#?★…【】《》？“”‘’！[\\]^_`{|}~]+'  # 不去除数字 不去除小数点
    # r = '[a-zA-Z’!"#$&\'*+<=>?@。#?★…【】《》？“”‘’！[\\]^_`{|}~]+'   # 把-去除掉了
    r = '[a-zA-Z]+'
    for j in range(firstContentC, columns):
        temp[firstContentR][j] = str(temp[firstContentR][j])
        for i in range(firstContentR + 1, firstDataR):
            # 如果单元格内容相同就不合并
            temp[i][j] = str(temp[i][j])
            # if (temp[i][j] != temp[i - 1][j]):
            #     # print(temp[i][j], temp[i - 1][j], temp[i][j] == temp[i - 1][j])
            #     temp[firstContentR][j] = temp[firstContentR][j] + (temp[i][j])
                # print(temp[firstContentR][j])
            temp[firstContentR][j] = temp[firstContentR][j] + (temp[i][j])
        # if type(temp[firstContentR][j]) != float:
        #     temp[firstContentR][j] = re.sub(r, '', temp[firstContentR][j]).replace('\n', '').replace(' ', '').replace('\u3000', '')
        # print(temp[firstContentR][j])
        temp[firstContentR][j] = re.sub(r, '', temp[firstContentR][j]).replace('\n', '').replace(' ', '').replace(
            '\u3000', '')
        # print(temp[firstContentR][j])
    # 写入excel
    for i in range(firstContentR, firstDataR):
        for j in range(firstContentC, columns):
            ws.write(i, j, temp[i][j])
    wb.save(targetfile)


# 针对一个文件夹下的所有excel文件进行合并单元格的处理，并且把处理完的文件存放在tergetdir中
def modifyDir(dir, targetdir):
    filelist = os.listdir(dir)
    xlslist = []
    tarxlslist = []
    if not os.path.exists(targetdir):
        os.mkdir(targetdir)
    for i in filelist:
        if i[-3:] == 'xls' or i[-4:] == 'xlsx':
            xlslist.append(dir + "\\" + i)
            tarxlslist.append(targetdir + "\\" + i)
    for i in range(len(xlslist)):
        modify(xlslist[i], tarxlslist[i])
    pifalingshou(targetdir)
    zhusucanyin(targetdir)


# 针对批发零售业表格的单独处理
def pifalingshou(dir):
    filelist = os.listdir(dir)
    xlslist = []
    for i in filelist:
        if i[-3:] == 'xls' or i[-4:] == 'xlsx':
            xlslist.append(dir + "\\" + i)
    for file in xlslist:
        rb = xlrd.open_workbook(file)
        table = rb.sheets()[0]
        if ("批" in table.cell_value(0, 0)) and ("零" in table.cell_value(0, 0)):
            print(file)
            columns = table.ncols
            rows = table.nrows
            wb = copy(rb)
            ws = wb.get_sheet(0)
            lis1 = []
            lis2 = []
            lisfinal = []
            for i in range(1, rows):
                if "一" in str(table.cell_value(i, 0)):
                    lis1.append(i)
                if "二" in str(table.cell_value(i, 0)):
                    lis2.append(i)
            for i in lis1:
                if "批发" in str(table.cell_value(i, 0)):
                    lisfinal.append(i)
                    break
            for i in lis2:
                if "零售" in str(table.cell_value(i, 0)):
                    lisfinal.append(i)
                    break
            if len(lisfinal) == 2:
                for i in range(lisfinal[0] + 1, lisfinal[1]):
                    ws.write(i, 0, "批发业" + str(table.cell_value(i, 0)))
                for i in range(lisfinal[1] + 1, rows):
                    ws.write(i, 0, "零售业" + str(table.cell_value(i, 0)))
            wb.save(file)


# 针对住宿餐饮的单独处理
def zhusucanyin(dir):
    filelist = os.listdir(dir)
    xlslist = []
    for i in filelist:
        if i[-3:] == 'xls' or i[-4:] == 'xlsx':
            xlslist.append(dir + "\\" + i)
    for file in xlslist:
        rb = xlrd.open_workbook(file)
        table = rb.sheets()[0]
        if ("住宿" in table.cell_value(0, 0)) and ("餐饮" in table.cell_value(0, 0)):
            print(file)
            columns = table.ncols
            rows = table.nrows
            wb = copy(rb)
            ws = wb.get_sheet(0)
            lis1 = []
            lis2 = []
            lisfinal = []
            for i in range(1, rows):
                if "一" in str(table.cell_value(i, 0)):
                    lis1.append(i)
                if "二" in str(table.cell_value(i, 0)):
                    lis2.append(i)
            for i in lis1:
                if "住宿" in str(table.cell_value(i, 0)):
                    lisfinal.append(i)
                    break
            for i in lis2:
                if "餐饮" in str(table.cell_value(i, 0)):
                    lisfinal.append(i)
                    break
            if len(lisfinal) == 2:
                for i in range(lisfinal[0] + 1, lisfinal[1]):
                    ws.write(i, 0, "住宿业" + str(table.cell_value(i, 0)))
                for i in range(lisfinal[1] + 1, rows):
                    ws.write(i, 0, "餐饮业" + str(table.cell_value(i, 0)))
            wb.save(file)

# pifalingshou(r'D:\project\energy data\全自动\无锡\2019\modify')
# modify(r'D:\project\energy data\全自动\无锡\2019\excel\10-02.xls', 'temp.xls')

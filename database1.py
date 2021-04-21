import pymysql
import datetime
import json
from PyQt5.QtWidgets import QMessageBox
from PyQt5 import QtGui
from ERRORS import NETWORK_ERR, INSERT_FAILURE, UPDATE_FAILURE, EXECUTE_FAILURE

# 数据库接口
class Database:
    def __init__(self):
        try:
            self.connect_to_database()
        except:
            raise NETWORK_ERR


    def connect_to_database(self):
        self.db = pymysql.connect(host="gz-cynosdbmysql-grp-0965sb99.sql.tencentcdb.com", user="root", password="A1b1c1d1", db="test", port=25462,charset='utf8')
        self.cr = self.db.cursor(cursor = pymysql.cursors.DictCursor)

    
    def add_data_excel(self, worksheet, year_list):
        """
        从excel表格中添加数据到数据库

        :param worksheet: excel表格
        :param year_list: 年份列表
        """
        json_body = []
        # 数据从excel的第2行开始
        for row_num in range(1, worksheet.nrows):
            row = worksheet.row_values(row_num)  # 读取excel的一行
           
            dict_field = {}
            dict_field['编码'] = str(int(row[0]))
            dict_field['地域'] = row[1]
            # 6个级别 从第2列开始
            # 尝试字符串化 (whh)
            for cat_index in range(2, 8):
                dict_field['级别%d' % (cat_index-1)] = str(row[cat_index])
                # dict_field['级别%d' % (cat_index-1)] = row[cat_index]
            # max_year-min_year个年份 从第8列开始
            tmp_year = {}
            for year_index in range(len(year_list)):
                try:
                    value = str(float(row[8 + year_index]))
                except:
                    value = '0'
                
                tmp_year[year_list[year_index]] = value
            
            jsn_year = json.dumps(tmp_year)
            dict_field['年份'] = jsn_year
          
            json_body.append(dict_field)
      
        sql = "INSERT INTO `Data` (`地域`, `编码`, `级别1`, `级别2`, `级别3`, `级别4`, `级别5`, `级别6`, `年份`) VALUES "
            
        for i in range(len(json_body)):
            dict_field = json_body[i]
            # print(dict_field)
            if i != 0:
                sub_str = ',('
            else:
                sub_str = '('
            
            sub_str += "'%s'," % dict_field["地域"]
            sub_str += "%d," % int(dict_field["编码"])
            for i in range(1, 7):
                ind = "级别%d" % i
                item = dict_field[ind]
                if isinstance(item, str):
                    sub_str += "'%s'" % item
                    sub_str += ","
                else:
                    sub_str += dict_field[ind]

            sub_str += "'%s'" % dict_field["年份"]
            sub_str += ")"
            sql += sub_str

        try:
            self.connect_to_database()
            self.cr.execute(sql)
            self.db.commit()
        except:
            raise INSERT_FAILURE()
      

    def add_data_dialog(self, data, years):
        """
        从对话框中获取信息并存入到数据库中
        :param data: 从对话框获得的数据
        :param years: 对话框中的表头包含的年份
        """
        dict_field = {}
        dict_field['编码'] = data[0]
        dict_field['地域'] = data[1]
        # 6个级别 从第2列开始
        for cat_index in range(2, 8):
            dict_field['级别%d' % (cat_index - 1)] = data[cat_index]
        # 从第8列开始
        tmp_year = {}
        for year_index in range(len(years)):
            value = data[8+year_index]
            tmp_year[years[year_index]] = value
            # dict_field[years[year_index]] = value

        jsn_year = json.dumps(tmp_year)
        dict_field['年份'] = jsn_year
        
        sql = "INSERT INTO `Data` (`地域`, `编码`, `级别1`, `级别2`, `级别3`, `级别4`, `级别5`, `级别6`, `年份`) VALUES ("
        sql += "'%s'," % dict_field["地域"]
        sql += "%d," % int(dict_field["编码"])
        for i in range(1, 7):
            ind = "级别%d" % i
            item = dict_field[ind]
            if isinstance(item, str):
                sql += "'%s'" % item
                sql += ","
            else:
                sql += dict_field[ind]
                sql += ","

        sql += "'%s'" % dict_field["年份"]
        sql += ")"

        try:
            self.connect_to_database()
            self.cr.execute(sql)
            self.db.commit()
        except:
            raise INSERT_FAILURE
        

    def modify_data(self, place, modified_data, header):
        """
        修改数据
        :param place: 修改过的数据在modified_data数组中的位置
        :param modified_data: 修改了部分数据的数组
        :param header: modified_data数据的键
        """
        query_list = []
        for item in place:
            region = modified_data[item[0]][1]
            index = modified_data[item[0]][0]
            if item[1] <= 7:
                modify_item = header[item[1]]
                sql = "UPDATE `Data` SET `%s`='%s' WHERE `地域`='%s' AND `编码`='%d'" % (modify_item, modified_data[item[0]][item[1]], region, int(index))

            else:
                query = "SELECT `年份` from `Data` WHERE `地域`='%s' AND `编码`='%d'" % (region, int(index))
                # print(query)
                self.cr.execute(query)
                result = self.cr.fetchall()[0]['年份']
                dic = json.loads(result)
                for i in range(8, len(header)):
                    dic[header[i]] = modified_data[item[0]][i]
                json_year = json.dumps(dic)
                sql = "UPDATE `Data` SET `年份`='%s' WHERE `地域`='%s' AND `编码`='%d'" % (json_year, region, int(index))
                
            query_list.append(sql)
        
        for query in query_list:
            try:
                self.connect_to_database()
                self.cr.execute(query)
                self.db.commit()
            except:
                raise UPDATE_FAILURE()


    def delete_data(self, index, region):
        """
        删除数据
        :param index: 所要删除的数据的编码
        :param region: 所要删除的数据所在的地域
        """
        query = "DELECT FROM `Data` WHERE `编码`='%d' AND `地域`='%s'" % (int(index), region)
        # query = 'delete from Data where "编码"=' + "'%s'" % index + 'and "地域"=' + "'%s'" % region
        try:
            self.cr.execute(query)
            self.db.commit()
        except:
            raise EXECUTE_FAILURE()


    def get_data(self, query):
        """
        从数据库获取数据
        :param query: 用于查询的sql语句
        """
        try:
            self.connect_to_database()
            self.cr.execute(query)
        except:
            raise EXECUTE_FAILURE()
            return

        result = self.cr.fetchall()

        for item in result:
            year = eval(item.pop("年份"))
            item = item.update(year)

        return result
        

    def empty_table(self):
        """
        清空数据库
        """
        sql = "truncate table `Data`"
        try:
            self.connect_to_database()
            self.cr.execute(sql)
            self.db.commit()
        except:
            raise EXECUTE_FAILURE()


    def retry_connect(self):
        """
        尝试重连数据库
        """
        try:
            self.db = pymysql.connect(host="gz-cynosdbmysql-grp-0965sb99.sql.tencentcdb.com", user="root", password="A1b1c1d1", db="test", port=25462,charset='utf8')
            self.cr = self.db.cursor(cursor = pymysql.cursors.DictCursor)
            return True
        except:
            return False


    def get_data_simple(self, query):
        try:
            self.connect_to_database()
            self.cr.execute(query)
        except:
            raise EXECUTE_FAILURE()
            return

        result = self.cr.fetchall()
        return result
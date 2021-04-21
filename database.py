from influxdb import InfluxDBClient
import win32api
import datetime



class Database:
    def __init__(self):
        win32api.ShellExecute(0, 'open', 'influxdb-1.8.2-1\\influxd.exe', '', '', 0)  # 运行influxd.exe
        self.client = InfluxDBClient('localhost', 8086, 'root', 'root', 'Energy')

        #  检测是否存在数据库Energy, 若否则创建
        flag = False
        for database in self.client.get_list_database():
            if database['name'] == 'Energy':
                flag = True
        if not flag:
            self.client.create_database('Energy')

    def add_data_excel(self, worksheet, year_list):
        query = 'select * from "Data"'
        points = self.get_data(query)
        temp = {}
        for point in points:
            temp['%s %s' % (point['编码'], point['地域'])] = point['time']

        temp_time = datetime.datetime.now()
        json_body = []
        # 数据从excel的第2行开始
        for row_num in range(1, worksheet.nrows):
            row = worksheet.row_values(row_num)  # 读取excel的一行
            dict_tag = {}
            dict_tag['编码'] = str(int(row[0]))
            dict_tag['地域'] = row[1]

            dict_field = {}
            # 6个级别 从第2列开始
            # 尝试字符串化 (whh)
            for cat_index in range(2, 8):
                dict_field['级别%d' % (cat_index-1)] = str(row[cat_index])
                # dict_field['级别%d' % (cat_index-1)] = row[cat_index]
            # max_year-min_year个年份 从第8列开始
            for year_index in range(len(year_list)):
                try:
                    value = str(float(row[8 + year_index]))
                except:
                    value = '0'
                dict_field[year_list[year_index]] = value

            dict_record = {}
            dict_record['tags'] = dict_tag
            dict_record['fields'] = dict_field
            dict_record['measurement'] = 'Data'

            try:
                dict_record['time'] = temp['%s %s' % (dict_tag['编码'], dict_tag['地域'])]
            except:
                temp_time += datetime.timedelta(microseconds=5)
                dict_record['time'] = temp_time.isoformat()

            json_body.append(dict_record)
        self.client.write_points(json_body)

    def add_data_dialog(self, data, years):
        query = 'select * from "Data"'
        points = self.get_data(query)
        temp = {}
        for point in points:
            temp['%s %s' % (point['编码'], point['地域'])] = point['time']

        temp_time = datetime.datetime.now()
        dict_tag = {}
        dict_tag['编码'] = data[0]
        dict_tag['地域'] = data[1]
        dict_field = {}
        # 6个级别 从第2列开始
        for cat_index in range(2, 8):
            dict_field['级别%d' % (cat_index - 1)] = data[cat_index]
        # 从第8列开始
        for year_index in range(len(years)):
            value = data[8+year_index]
            dict_field[years[year_index]] = value

        dict_record = {}
        dict_record['tags'] = dict_tag
        dict_record['fields'] = dict_field
        dict_record['measurement'] = 'Data'

        try:
            dict_record['time'] = temp['%s %s' % (dict_tag['编码'], dict_tag['地域'])]
        except:
            dict_record['time'] = temp_time.isoformat()

        self.client.write_points([dict_record])

    def delete_data(self, index, region):
        query = 'delete from Data where "编码"=' + "'%s'" % index + 'and "地域"=' + "'%s'" % region
        result = self.client.query(query)

    def get_data(self, query):
        result = self.client.query(query)
        points = list(result.get_points())
        print(points)
        return points



if __name__ == '__main__':
    db = Database()
    query = 'select * from "Data"'
    db.get_data(query)
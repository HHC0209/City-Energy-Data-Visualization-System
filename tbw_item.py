# 定义表格中一个数据项的类
class tbw_item(object):
    def __init__(self):
        super().__init__()
        self.name = []
        self.index = []
        self.data = []
        self.year = []
        self.row = 0
        self.realname = ''

    def getName(self):
        return self.name
    
    def setName(self, name):
        self.name = name

    def getIndex(self):
        return self.index

    def setIndex(self, index):
        self.index = index

    def setRow(self, row):
        self.row = row

    def getRow(self, row):
        return self.row

    def setData(self, data):
        self.data = data

    def getData(self):
        return self.data

    def setRealName(self):
        region = self.data[0]['地域']
        self.realname = region
        for item in self.name:
            self.realname += ' > '
            self.realname += item
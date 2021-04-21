# 自定义各种错误
class NETWORK_ERR(Exception):
    def __init__(self):
        self.error_message = "无法连接网络。"
    
    def __str__(self):
        return self.error_message


class INSERT_FAILURE(Exception):
    def __init__(self):
        self.error_message = "导入数据失败，请检查网络连接，请勿使用已存在的编码"

    def __str__(self):
        return self.error_message

class UPDATE_FAILURE(Exception):
    def __init__(self):
        self.error_message = "更新数据失败"

    def __str__(self):
        return self.error_message

class EXECUTE_FAILURE(Exception):
    def __init__(self):
        self.error_message = "无法完成操作，请检查网络"
    def __str__(self):
        return self.error_message
from openpyxl import Workbook

class Sheet:
    def __init__(self, worksheet):
        self.__sheet = worksheet

    def __str__(self):
        # 使用列表推导式构建一个包含所有行的字符串列表
        rows = ['\t'.join(map(str, row)) for row in self.__sheet.iter_rows(values_only=True)]
        # 使用换行符连接所有行，构建一个完整的字符串
        return '\n'.join(rows)

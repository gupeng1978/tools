from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

class Sheet:
    def __init__(self, worksheet, name='Sheet1'):
        if not isinstance(worksheet, Worksheet):
            raise TypeError("worksheet must be an instance of openpyxl.worksheet.worksheet.Worksheet")

        self.__sheet = worksheet
        self.__name = name


    def __str__(self):
        # 添加工作表名称
        sheet_name_info = f"------- Sheet Name: {self.__name}"

        # 使用列表推导式构建一个包含所有行的字符串列表
        rows = ['\t'.join(map(str, row)) for row in self.__sheet.iter_rows(values_only=True)]

        # 添加合并单元格的信息
        merged_cells_info = "Merged Cells:"
        for merged_range in self.__sheet.merged_cells.ranges:
            merged_cells_info += f"\n{merged_range.coord}"

        # 使用换行符连接工作表名称、所有行和合并单元格的信息，构建一个完整的字符串
        return sheet_name_info + '\n' + '\n'.join(rows) + '\n' + merged_cells_info

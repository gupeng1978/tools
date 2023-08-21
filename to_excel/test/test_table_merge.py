import unittest
from openpyxl import Workbook
import sys
sys.path.append('../table')
sys.path.append('../sheet')

from table import Table_Merge
from sheet import Sheet


class TestTableMerge(unittest.TestCase):

    def test_merge_header(self):
        ws = Workbook().active
        sheet = Sheet(ws)
        ws.append(["A",     "A",    "B",    "C",    "C"])
        ws.append(["A1",    "A2",   "B1",   "C1",   "C2"])
        table_info = {'header': {'row_start':1, 'row_end' : 2},}
        sort_key_row_index = [0,] # 根据A1排序并合并
        table_merge = Table_Merge(ws, table_info, sort_key_row_index)
        
        print('----before merge --------')
        print(sheet)
        
        table_merge.merge()

        print('----after merge --------')
        print(sheet)
        
        
        # 使用断言方法验证合并后的单元格
        # self.assertEqual(ws.merged_cells.ranges[0].coord, "A1:B1")
        # self.assertEqual(ws.merged_cells.ranges[1].coord, "C2:D2")

# 如果是直接运行这个文件，那么执行测试
if __name__ == '__main__':
    unittest.main()

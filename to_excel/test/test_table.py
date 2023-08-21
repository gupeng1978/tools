import unittest
from openpyxl import Workbook
import sys
sys.path.append('../table')
sys.path.append('../sheet')

from table import Header, Record, Table
from sheet import Sheet


class TestTable(unittest.TestCase):
    
    def test_table_case01(self):
        header = Header()
        header.add(["model", "thread", "step","device_nn", "device_nn", "host_pcie_nn", "host_pcie_nn"])
        header.add(["model", "thread", "step","dev_ms",    "dev_fps",   "host_ms",      "host_fps"])
        header.set_active(1, ['model','thread'])
        header.set_alias({"model": "模型", 
                          "thread": "线程数", 
                          "step": "步骤", 
                          "dev_ms": "时间(ms)", 
                          "dev_fps": "帧率", 
                          "host_ms": "时间(ms)", 
                          "host_fps": "帧率",
                          "device_nn":"device推理",
                          "host_pcie_nn":"Host PCIe推理"})
        record = Record()
        
        record.add_from_dict({"model": "resnet50", 
                              "thread": '2', 
                              "step" :  "dclmdlLoadFromFile",
                              "dev_ms" : "2",
                              "host_ms" : "1.59"})

        record.add_from_dict({"model": "resnet50", 
                              "thread": '1', 
                              "step" :  "dclmdlLoadFromFile",
                              "dev_ms" : "2.59",
                              "host_ms" : "3.59"})

        record.add_from_dict({"model": "resnet50", 
                              "thread": '1', 
                              "step" :  "copyH2D",
                              "host_ms" : "0.5"})

        record.add_from_dict({"model": "resnet50", 
                              "thread": '1', 
                              "step" :  "copyD2H",
                              "host_ms" : "1.5"})

        record.add_from_dict({"model": "yolov8", 
                              "thread": '16', 
                              "step" :  "copyD2H",
                              "host_ms" : "3.5"})

        record.add_from_dict({"model": "resnet50", 
                              "thread": '2', 
                              "step" :  "copyH2D",
                              "host_ms" : "0.5"})

        record.add_from_dict({"model": "resnet50", 
                              "thread": '4', 
                              "step" :  "copyD2H",
                              "host_ms" : "3.5"})

        print(record)

         # 创建一个工作簿和工作表
        workbook = Workbook()
        worksheet = workbook.active

        table = Table(worksheet, header, record)
        table.merge_cells()
        table.set_attrs()
        
        
                
        print(table)        
        print(Sheet(worksheet,"nnp perf"))

         # 保存工作簿
        workbook.save("nnp-perf-case01.xlsx")
        
        
        

   
    

# 如果是直接运行这个文件，那么执行测试
if __name__ == '__main__':
    unittest.main()

import unittest
from ptable import gen_excel_table, get_default_config

import os
TEST_CONFIG_file = os.path.join(os.path.dirname(__file__), '../data/test_cfg.yaml')
DEMO_CONFIG_file = os.path.join(os.path.dirname(__file__), '../data/demo_cfg.yaml')
DEMO_LOG_FILE = os.path.join(os.path.dirname(__file__), '../data/demo_log.txt')

class TestSheet(unittest.TestCase):
    
    def test_sheet_case01(self):
        if os.path.isfile(TEST_CONFIG_file):
            os.remove(TEST_CONFIG_file)
        get_default_config(TEST_CONFIG_file) 
        self.assertEqual(os.path.isfile(TEST_CONFIG_file), True)
        pass
    
    def test_sheet_case02(self):
        gen_excel_table(DEMO_CONFIG_file)
        pass
 

# 如果是直接运行这个文件，那么执行测试
if __name__ == '__main__':
    unittest.main()

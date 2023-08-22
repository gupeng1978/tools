import unittest

# 定义一个测试加载器
loader = unittest.TestLoader()

# 从'test'目录中发现并加载所有测试
suite = loader.discover('test')

# 运行所有测试
runner = unittest.TextTestRunner()
runner.run(suite)

from setuptools import setup, find_packages
from ptable import __version__


print(__version__)

# 读取README.md文件内容
with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()
    
setup(
    name='xlsx_ptable',
    version=__version__,
    packages=find_packages(),
    install_requires=[
        # 列出项目的依赖包，例如：
        'openpyxl==3.1.2',
        'PyYAML',
    ],
    author='Peng Gu',
    author_email='gu.peng@intellif.com',
    description='Extract information from log files and create Excel spreadsheets.',
    long_description=long_description,
    long_description_content_type="text/markdown",
    url='https://github.com/gupeng1978/tools',
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
    ],
)

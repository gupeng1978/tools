from setuptools import setup, find_packages

setup(
    name='xlsx_ptable',
    version='0.1.0',
    packages=find_packages(),
    install_requires=[
        # 列出项目的依赖包，例如：
        'requests',
    ],
    author='Peng Gu',
    author_email='gu.peng@intellif.com',
    description='Extract information from log files and create Excel spreadsheets.',
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

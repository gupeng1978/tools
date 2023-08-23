# xlsx_ptable python包介绍

xlsx_ptable Python Package是一个基于用户log文件，并根据配置文件生成excel表格的python库。它可以帮助你生成优雅的Excel表格。

## 安装

使用以下命令安装My Python Package：

```bash
python -m pip install xlsx_ptable
```

## 用法
以下是如何使用My Python Package的一些基本示例：

假设用户user_log.txt文件为:

```
%tab1: (X1@1), (X2@a), (Y@sample y1), (Z@sample z1)
.............
%tab1: (X1@2), (X2@), (Y@sample y2), (Z@sample z2)
.............
%tab1: (X1@3), (Z@sample z3)
%tab1: (X1@4), (X2@c), (Y@sample y3), (Z@sample z4)
%tab1: (X1@4), (X2@d), (Y@sample y4),
%tab1: (X1@5), (X2@e), (Y@sample y5), (Z@sample z5)
.............
%tab1: (X1@5), (X2@f), (Y@sample y6), (Z@sample z6)
```
用户配置文件user.cfg(YAML格式)为：
```yaml
excel_path: 'sample1.xlsx' # 输出的excel表格路径
sheets:
  tag: [sheet1] #excel表格sheet名称
tables:
  # 表1
  - sheet_tag: sheet1 # 这个表属于sheet1
    name: "%tab1" # 该表的records记录在path/to/table1.log中以%tab1为关键字的行中
    head-0: [X,  X, Y, Z]   #一级表头(不支持中文)
    head-1: [X1, X2, Y, Z] #二级表头(不支持中文), 长度必须和一级表头相等
    head-key: [X1] # 列方向上以X1排序且合并
    alias: # 表头的别名，支持中文转译
      X: 表头1
      Y: 表头2
      Z: 表头3
      X-1: 选择1
      X-2: 选择2
    record_file: 'user_log.txt' # log记录，每行最多一个record记录，格式为 ...%tab1...(X1@2.0)...(X2@3)...(Y@content1)...(Z@content2)
```

使用方法如下：
```python
from ptable import gen_excel_table
gen_excel_table("user.cfg") 
```

输出excel表格为:

| 表头1 |           |表头2    | 表头3    |
|-------|----------|----------|----------|
| X1    | X2       |          |
| 1     | a        | sample y1 | sample z1 |
| 2     |          | sample y2 | sample z2 |
| 3     |          |           | sample z3 |
| 4     | c        | sample y3 | sample z4 |
|       | d        | sample y4 |           |
| 5     | e        | sample y5 | sample z5 |
|       | f        | sample y6 | sample z6 |


## 注意事项
### log的格式要求
1. 表格中每个记录只能在log中一行
2. 必须携带表格tag标记(比如例子中%tab1)
3. 每个记录由(key@value)表示一个cell的记录，其中key为表头head-1中的关键字
4. 对于head-key关键字，记录中必须包含head-key的关键字，比如例子中head-key为[X1]，那么记录必须有(X1@value)的单元值

### 配置文件要求
1. 符合YAML配置文件规范
    1. 保留字符（例如%，！#，：，!等）应用引号引起来。
    2. 确保缩进一致，使用空格而不是制表符。
    3. 列表项应以短划线和空格开始。
2. 支持多sheet多表格, 由于每个table会动态计算列的宽度并调整sheet的列宽度，所以一般一个sheet中存放一个table;
3. table中所在的sheet_tag，必须在sheets的tag中；
4. table支持2级header表示，其中head-1中包括记录log中所允许的keys；
5. table会动态在head进行行方向、列方向合并；
6. alias为header的别名，主要解决log中都是英文记录，而表格header需要中文解释，如果记录的key太长，也可以通过alias别名替代(比如log中使用$0代表"dclxxxyyyzzz")
7. **head-key**为关键key，所有表格记录会基于该key，在列方向上排序以及合并;
8. 支持单个log存放多个table 记录;


## Sample
更多sample，请访问[xlsx_ptable](https://github.com/gupeng1978/tools/tree/main/xlsx_ptable/sample)。

## 贡献
欢迎任何形式的贡献！请阅读我们的贡献指南了解如何开始。

##  许可
根据MIT许可证发布。

## 问题和建议
如果你有任何问题或建议，请通过邮箱gu.peng@intelllif.com联系或者
请[在此提出Issue](https://github.com/gupeng1978/tools/issues)。

'''
    这是一个用于拆分xlsx的项目
    1. 输入指定文件，指定条件用于拆分
    2. 拆分的文件存于当前目录下的out目录
'''

import os

import pandas as pd

# file_xlsx：文件位置
file_xlsx = input('文件目录及文件名（完整路径）：')
# file_link：拆分依据
file_link = input('需要拆分的列名（完全一致）：')
# out_dir：输出目录
out_dir = input('输出的目录（请确认已经建立）：')

# 验证目录是否存在
if os.system(f'mkdir {out_dir}'):
    print('目录已存在！')
else:
    print('目录已建立！')
    

# 引入dataframe格式
df = pd.read_excel(file_xlsx, index_col=0)

# link_list：选择列的不重复列表
link_list = list(dict.fromkeys(df[file_link]))

# new_file：拆分的dataframe导入列表
new_file = []

# 按照'目录'列拆分并写入
for i in link_list:
    out = os.path.join(out_dir, i)
    df[df[file_link] == i].to_excel(f'{out}.xlsx')
    print(f'{i}.xlsx文件已经写入！')

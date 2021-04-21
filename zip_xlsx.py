'''
    项目：合并目录下的xlsx或者xls文件
    需要的包：pandas， xlrd。
    以上，个人风格，欢迎指正，看心情改正。
'''
# 导入os和pandas包
import os
import pandas as pd 

# files：目录下的文件
files = []

# 遍历文件夹下所有xls或者xlsx文件并列入files集合中
for dirs, _, names in os.walk(input('输入要合并的目录：')):
    for name in names:
        if os.path.splitext(name)[1] in ['.xlsx', '.xls']:
            files.append(os.path.join(dirs, name))

# df：输出到文件用的变量
df = pd.DataFrame()

# head_link：标题栏
head_link = int(input('标题栏最后一格：'))-1

# 遍历files，把内容叠加到df中
for i in files:
    df = df.append(pd.read_excel(i, header=head_link, index_col=0))

# 输出到文件
try:
    save_file = input('输入保存地址及文件名：')
    df.to_excel(save_file)
    print(f'{save_file}输出完成！')
except:
    print('输出错误！')
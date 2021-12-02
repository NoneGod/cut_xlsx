'''
    这是一个用于合并拆分xlsx的项目
    也就是把之前的两个功能合并，并加入指引，引入面向对象的概念，之后可能考虑做个web的界面
    作者：NoCode
    版本：1.04
'''

import os
# from posixpath import splitext
import pandas as pd

# 字段名
def colnum_list(file_xlsx, file_header):
    engine = 'xlrd'
    if os.path.splitext(file_xlsx)[1] == '.xlsx':
        engine = 'openpyxl'
    return list(dict.fromkeys(pd.read_excel(file_xlsx, header=file_header, index_col=0, engine=engine)))

# 合并
def zip_excel(zip_dir:str, zip_link:int, save_file:str):
    # files：目录下的文件
    files = []

    # 遍历文件夹下所有xls或者xlsx文件并列入files集合中
    for dirs, _, names in os.walk(zip_dir):
        for name in names:
            if os.path.splitext(name)[1] in ['.xlsx', '.xls']:
                files.append(os.path.join(dirs, name))

    # df：输出到文件用的变量
    df = pd.DataFrame()

    # head_link：标题栏
    head_link = int(zip_link)-1

    # 遍历files，把内容叠加到df中
    for i in files:
        engine = 'xlrd'
        if os.path.splitext(i)[1] == '.xlsx':
            engine = 'openpyxl'
        df = df.append(pd.read_excel(i, header=head_link, engine=engine))  # index_col=0, 索引列为当前列

    # 输出到文件
    try:
        df.to_excel(save_file)
        print('已输出到：{}'.format(save_file))
    except:
        print('输出错误！')

# 拆分
def cut_excel(file_xlsx:str, file_link:str, file_header:int, out_dir:str):   # file_xlsx：文件位置    file_link：拆分依据   file_header：表头   out_dir：输出目录
    # 按照文件名建立目录
    if os.system('mkdir {}'.format(out_dir)):
        print('已经建立：{}'.format(out_dir))
    else:
        print('{}\t目录已经存在！'.format(out_dir))
    
    # 导入文件成pandas格式，并取消首列索引
    engine = 'xlrd'
    if os.path.splitext(file_xlsx)[1] == '.xlsx':
        engine = 'openpyxl'
    data_file = pd.read_excel(file_xlsx, header=file_header, index_col=0, engine=engine)

    # colnum_data:目录列列表，生成目标列去重列表
    colnum_data = list(dict.fromkeys(data_file[file_link]))

    # 按照'目录'列拆分并写入
    for col_data in colnum_data:
        out_name = os.path.join(out_dir, col_data)
        data_file[data_file[file_link] == col_data].to_excel('{}.xlsx'.format(out_name))
        print('已经拆分:{}.xlsx'.format(out_name))

# 程序入口
if __name__ == '__main__':
    flag = True
    count = True
    while flag == True:
        if count == True:
            count = False
        else:
            os.system('pause')
        os.system('cls')
        print('{}\n'.format('-'*50))
        print('''
        作者：NoCode， 版本：1.04
        输入数字：
        1.合并
        2.拆分
        0.退出
        ''')

        print('{}\n'.format('-'*50))
        while True:
            choose_number = int(input('输入数字选择：'))

            # 合并
            if choose_number == 1:
                zip_dir = input('文件目录（默认当前目录）：')
                if zip_dir == '':
                    zip_dir = os.path.join(os.getcwd())

                save_file = input('输出合并文件位置及文件名(默认当前目录下创建out.xlsx)：')
                if save_file == '':
                    save_file = os.path.join(os.getcwd(), 'out.xlsx')

                zip_link = input('标题行数(默认1)：')
                if zip_link == '':
                    zip_link = 1
                else:
                    try:
                        zip_link = int(zip_link)
                    except Exception as e:
                        print('发现错误：{}。'.format(e))
                        zip_link = 1
                zip_excel(zip_dir, zip_link, save_file)
                break

            # 拆分
            elif choose_number == 2:
                file_header = 0
                file_xlsx = input('需要拆分的文件（默认当前目录下的out.xlsx）：')
                if file_xlsx == '':
                    if os.path.isfile('out.xlsx'):
                        file_xlsx = os.path.join(os.getcwd(), 'out.xlsx')
                elif os.path.splitext(file_xlsx)[1] in ['.xlsx', '.xls']:
                    file_xlsx = os.path.join(os.getcwd(), file_xlsx)
                    file_header = input('输入标题栏的位置（默认1）：')
                    if file_header != '':
                        file_header = int(file_header)-1
                else:
                    print('输入错误！')
                    break
                out_dir = input('保存的位置（默认当前目录下创建out目录）：')
                if out_dir == '':
                    out_dir = os.path.join(os.getcwd(), 'out')
                else:
                    out_dir = os.path.join(os.getcwd(), out_dir)
                
                # 显示字段
                list_data = dict(zip(range(len(colnum_list(file_xlsx, file_header))), colnum_list(file_xlsx, file_header)))
                for i in range(len(colnum_list(file_xlsx, file_header))):
                    print('{}\t{}'.format(i, list_data[i]))       

                # 字段
                file_link = input('选择需要拆分的字段前面的数字：')
                file_link = list_data[int(file_link)]
                # print(file_link)
                try:
                    cut_excel(file_xlsx, file_link, file_header, out_dir)
                    print('拆分已完成！')
                    break
                except Exception as e:
                    print('出现错误：{}'.format(e))
                    break
            else:
                flag = False
                break

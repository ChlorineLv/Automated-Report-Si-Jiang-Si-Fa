# -*- coding: UTF-8 -*-
'''
### Version: Python 3.7.4
### Date: 2019-11-11 14:56:42
### LastEditors: ChlorineLv@outlook.com
### LastEditTime: 2019-11-12 11:37:49
### Description: 统计各人装移机单数
'''

import os
import pandas as pd 
import time

if __name__ == "__main__":
    csv_file = f'{os.path.dirname(__file__)}\表1装移机工单一览表.csv'
    csv_file = input(f'请输入装移机工单一览表文件名：（默认：{csv_file}）\n') or csv_file
    # 下载下来的csv像是gbk格式
    df = pd.DataFrame(pd.read_csv(csv_file, encoding='gbk'))
    # 取反，取施工类型不包括拆机
    df = df[~df['施工类型'].isin(['拆机'])]
    # 取反，取施工类型不包括其他
    df = df[~df['施工类型'].isin(['其他'])]
    # 取反，取施工类型不包括移拆机
    df = df[~df['施工类型'].isin(['移拆'])]
    # 以处理人工号计算频次
    df_final = df.groupby(by='处理人工号').size()
    # 删掉不需要使用的df，节省内存
    del df
    print(df_final)
    temp_excel_file = f'{os.path.dirname(__file__)}\中间表：人员装移机名单{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    df_final.to_excel(excel_writer = temp_excel_file, index = True)
    
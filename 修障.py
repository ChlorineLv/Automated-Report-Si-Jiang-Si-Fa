# -*- coding: UTF-8 -*-
'''
### Version: Python 3.7.4
### Date: 2019-11-12 17:37:05
### LastEditors: ChlorineLv@outlook.com
### LastEditTime: 2019-11-12 17:56:08
### Description: 统计各人修账单数，去除部门分局“合作方（甘肃万维）”“合作方（天讯）”
'''

import os
import pandas as pd 
import time

if __name__ == "__main__":
    csv_file = f'{os.path.dirname(__file__)}\表2障碍用户申告一览表.csv'
    csv_file = input(f'请输入表2障碍用户申告一览表文件名：（默认：{csv_file}）\n') or csv_file
    # 下载下来的csv像是gbk格式
    df = pd.DataFrame(pd.read_csv(csv_file, encoding='gbk'))
    # 取反，取部门分局不包括合作方（甘肃万维）
    df = df[~df['部门分局'].isin(['合作方（甘肃万维）'])]
    # 取反，取部门分局不包括合作方（天讯）
    df = df[~df['部门分局'].isin(['合作方（天讯）'])]
    # 以处理人工号计算频次
    df_final = df.groupby(by='修理员工号').size()
    # 删掉不需要使用的df，节省内存
    del df
    print(df_final)
    temp_excel_file = f'{os.path.dirname(__file__)}\中间表：人员修障名单{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    df_final.to_excel(excel_writer = temp_excel_file, index = True)
    
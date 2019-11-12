# -*- coding: UTF-8 -*-
'''
### Version: Python 3.7.4
### Date: 2019-11-07 09:45:59
### LastEditors: ChlorineLv@outlook.com
### LastEditTime: 2019-11-12 10:17:25
### Description: 筛选人员名单
'''

import os
import pandas as pd
import time


if __name__ == "__main__":
    excel_file = f'{os.path.dirname(__file__)}\广州电信分公司_用户管理导出(2019-11-04).xlsx'
    excel_file = input("请输入文件名：(默认：{})\n".format(excel_file)).strip() or excel_file
    df = pd.DataFrame(pd.read_excel(excel_file))[['综调登录工号', '姓名', '手机', '分公司', '单位名称', '人员类别', '岗位类型', '在职状态']]
    df = df.loc[df['岗位类型'] == '外线施工岗'].loc[df['在职状态'] == '在职']
    df['手机'] = df['手机'].astype('int64')
    print(df)
    temp_excel_file = f'中间表人员名单{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    df.to_excel(excel_writer = temp_excel_file, index = False)
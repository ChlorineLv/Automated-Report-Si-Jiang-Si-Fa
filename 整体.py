# -*- coding: UTF-8 -*-
'''
### Version: Python 3.7.4
### Date: 2019-11-14 09:53:16
### LastEditors: ChlorineLv@outlook.com
### LastEditTime: 2019-11-14 17:11:34
### Description: 
'''

import os
import pandas as pd
import time

def get_name_list():
    '''
    ### description: 
    ### param {type} 
    ### return: 以工号为key的人员名单dict
    ### example: {'工号1': {'姓名': '1', '手机': 1, '分公司': '1', ...}, '工号2': {'姓名': '2', '手机': 2, '分公司': '2', ...}, ......}
    '''
    excel_file = f'{os.path.dirname(__file__)}\广州电信分公司_用户管理导出(2019-11-04).xlsx'
    excel_file = input("请输入文件名：(默认：{})\n".format(excel_file)).strip() or excel_file
    # 只保留下面的列
    df = pd.DataFrame(pd.read_excel(excel_file))[['综调登录工号', '姓名', '手机', '分公司', '单位名称', '人员类别', '岗位类型', '在职状态']]
    # 字段内容筛选
    df = df.loc[df['岗位类型'] == '外线施工岗'].loc[df['在职状态'] == '在职']
    # 将手机号完整记录，不再是科学记数法
    df['手机'] = df['手机'].astype('int64')
    # 工号列变index
    df.set_index('综调登录工号', inplace=True)
    # print(df)
    # 以index作为dict的key
    return df.to_dict(orient='index')


def get_zhuang_list():
    '''
    ### description: 
    ### param {type} 
    ### return: 以工号为key的各人装机次数dict
    ### example: {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    csv_file = f'{os.path.dirname(__file__)}\表1装移机工单一览表.csv'
    csv_file = input(f'请输入装移机工单一览表文件名：（默认：{csv_file}）\n') or csv_file
    # 下载下来的csv像是gbk格式
    df = pd.DataFrame(pd.read_csv(csv_file, encoding='gbk'))
    # 取反，取施工类型不包括拆机、其他、移拆机
    df = df[~df['施工类型'].isin(['拆机'])]
    df = df[~df['施工类型'].isin(['其他'])]
    df = df[~df['施工类型'].isin(['移拆'])]
    # 以处理人工号计算频次
    df = df.groupby(by='处理人工号').size()
    # 先转为dataframe再转为dict
    return df.to_frame().to_dict(orient='index')


def get_xiu_list():
    '''
    ### description: 
    ### param {type} 
    ### return: 
    ### example: 
    '''
    csv_file = f'{os.path.dirname(__file__)}\表2障碍用户申告一览表.csv'
    csv_file = input(f'请输入表2障碍用户申告一览表文件名：（默认：{csv_file}）\n') or csv_file
    # 下载下来的csv像是gbk格式
    df = pd.DataFrame(pd.read_csv(csv_file, encoding='gbk'))
    # 取反，取部门分局不包括合作方（甘肃万维）
    df = df[~df['部门分局'].isin(['合作方（甘肃万维）'])]
    # 取反，取部门分局不包括合作方（天讯）
    df = df[~df['部门分局'].isin(['合作方（天讯）'])]
    # 以处理人工号计算频次
    df = df.groupby(by='修理员工号').size()
    # 先转为dataframe再转为dict
    return df.to_frame().to_dict(orient='index')


if __name__ == "__main__":
    # 获取名单
    dict_name = get_name_list()
    # 获取装移工作量
    dict_zhuang = get_zhuang_list()
    # 获取修障工作量
    dict_xiu = get_xiu_list()
    for i in dict_name:
        dict_name[i]['装移机'] = dict_zhuang.get(i, {0:0})[0]
        dict_name[i]['修障'] = dict_xiu.get(i, {0:0})[0]
        dict_name[i]['光衰整治'] = 0
        dict_name[i]['合计'] = dict_name[i]['装移机'] + dict_name[i]['修障'] + dict_name[i]['光衰整治']
        dict_name[i]["dict_name[i]['光衰整治']"] = 0 if dict_name[i]['合计']/20 < 8 else 1
        
    df = pd.DataFrame.from_dict(dict_name, orient='index')
    temp_excel_file = f'{os.path.dirname(__file__)}\人员产能表：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    df.to_excel(excel_writer = temp_excel_file, index = True)
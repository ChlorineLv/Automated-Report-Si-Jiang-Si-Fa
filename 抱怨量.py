# -*- coding: UTF-8 -*-
'''
### Version: Python 3.7.4
### Date: 2019-11-18 16:50:21
### LastEditors: ChlorineLv@outlook.com
### LastEditTime: 2019-11-19 09:39:00
### Description: 计算抱怨量
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
    # excel_file = input(f'请输入《人员名单》文件名：(默认：{excel_file})\n').strip() or excel_file
    print(f'正在处理:{excel_file}')
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


def get_bao_list(excel_file, excel_sheet, column_name, column_judge):
    '''
    ### description: 以工号为key的各人抱怨/绿通次数
    ### param { (str)excel文件所在位置, (str)sheet名称, (str)用于筛选的列名, (str)工号列名 } 
    ### return: dict 
    ### example: {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    # excel_file = input(f'请输入《抱怨清单统计2019年10月》文件名：（默认：{excel_file}）\n').strip() or excel_file
    print(f'正在处理:{excel_file}')
    # 下载下来的excel像是gbk格式
    df = pd.DataFrame(pd.read_excel(excel_file, sheet_name=excel_sheet))[[column_name, column_judge]]
    # 以处理人工号计算频次
    df = df.groupby(by=column_name).size()
    # 先转为dataframe再转为dict
    return df.to_frame().to_dict(orient='index')


if __name__ == "__main__":
    # 获取名单
    dict_name = get_name_list()
    # 获取抱怨量
    file_baoyuan = f'{os.path.dirname(__file__)}\抱怨清单统计2019年10月.xlsx'
    dict_baoyuan = get_bao_list(file_baoyuan, '清单', '工号', '服务类别')
    file_lvtong = f'{os.path.dirname(__file__)}\绿通单清单2019年10月.xlsx'
    dict_lvtong = get_bao_list(file_lvtong, '1', '处理人工号', '服务类型')
    n = 0
    for i in dict_name:
        i1 = i[1:] if i.startswith('0') else i
        dict_name[i]['抱怨'] = dict_baoyuan.get(i1, {0:0})[0] + dict_lvtong.get(i1, {0:0})[0]
        
        
        n+=1
                
    df = pd.DataFrame.from_dict(dict_name, orient='index')
    print(df)
    temp_excel_file = f'{os.path.dirname(__file__)}\中间表：人员产能表：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    df.to_excel(excel_writer = temp_excel_file, index = True)
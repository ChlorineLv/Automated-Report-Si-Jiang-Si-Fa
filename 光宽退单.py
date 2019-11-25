# -*- coding: UTF-8 -*-
'''
:Version: Python 3.7.4
:Date: 2019-11-25 10:10:18
:LastEditors: ChlorineLv@outlook.com
:LastEditTime: 2019-11-25 10:43:53
:Description: 光宽退单
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


def get_specify_df(excel_file, excel_sheet, column_list):
    '''
    :description: 返回保留某几列的dataframe
    :param excel_file {str} : excel文件所在位置
    :param excel_sheet {str} : sheet名称
    :param column_list {list} : 工号等列名
    :return: dataframe
    '''
    print(f'正在处理:{excel_file}')
    df = pd.DataFrame(pd.read_excel(excel_file, sheet_name=excel_sheet))[column_list]
    return df


def specify_df_frequency(df_get, c_name, c_judge):
    '''
    :description: 返回df中各项含特定字段的频次
    :param df_get {dataframe} : dataframe
    :param c_name {str} : 需要保留的主要字段，如工号
    :param c_judge {dict} : 用于判断的列:值对，如{'服务3':'超时未修复故障'}
    :return: dataframe     [处理人工号  0]
                           [工号1,     1]
                           [工号2,     1]
                           [工号3,     1]
                           [工号4,     1]
    '''
    for (k,v) in c_judge.items():
        key = k
        value = c_judge[k]
    df_get[c_name] = df_get[c_name].apply(str)
    df = df_get.loc[df_get[key]==value][[c_name, key]].groupby(by=c_name).count()
    df.columns = [0]
    # temp_excel_file = f'{os.path.dirname(__file__)}\中间表：{c_name}{value}：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    # df.to_excel(excel_writer = temp_excel_file, index = True)
    return df


if __name__ == "__main__":
    t_start = time.time()
    """ 获取名单 """
    dict_name = get_name_list()
    """ 获取退单dict """
    file_tuidan = f'{os.path.dirname(__file__)}\\10月光宽退单规范.xlsx'
    # file_tuidan = input(f'请输入《10月光宽退单规范》文件名：（默认：{file_tuidan}）\n').strip() or file_tuidan
    df_tuidan = get_specify_df(file_tuidan, 'Sheet2', ['工号', '退单是否规范'])
    dict_tuidan = specify_df_frequency(df_tuidan, '工号', {'退单是否规范': '不规范'}).to_dict(orient='index')
    print(dict_tuidan)
    n = 0
    for i in dict_name:
        """ 《催装催修》拿过来时少了个0 """
        if n == 0:
            print('已为《催装催修》中头部缺少0的工号进行匹配前的适配……')
        i1 = i[1:] if i.startswith('0') else i
        dict_name[i]['光宽缓装虚假退单'] = dict_tuidan.get(i1, {0:0})[0]
        
        n+=1
                
    df = pd.DataFrame.from_dict(dict_name, orient='index')
    # print(df)
    temp_excel_file = f'{os.path.dirname(__file__)}\中间表：人员抱怨无理失约：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    df.to_excel(excel_writer = temp_excel_file, index = True)
    print(f'已完成，保存地址{temp_excel_file}\n总耗时{time.time() - t_start}秒')
    
# -*- coding: UTF-8 -*-
'''
:Version: Python 3.7.4
:Date: 2019-11-25 14:29:00
:LastEditors: ChlorineLv@outlook.com
:LastEditTime: 2019-11-25 16:21:08
:Description: 工信部
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
    df = pd.DataFrame(pd.read_excel(excel_file, sheet_name=excel_sheet, dtype={column_list[0]: str}))[column_list]
    return df


def specify_df_frequency(df_get, c_name, c_judge=0):
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
    df_get[c_name] = df_get[c_name].apply(str)
    if c_judge == 0:
        df = df_get.groupby(by=c_name).count()
    else:
        for (k,v) in c_judge.items():
            key = k
            value = c_judge[k]
        df = df_get.loc[df_get[key]==value][[c_name, key]].groupby(by=c_name).count()
    df.columns = [0]
    # temp_excel_file = f'{os.path.dirname(__file__)}\中间表：{c_name}{value}：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    # df.to_excel(excel_writer = temp_excel_file, index = True)
    return df


if __name__ == "__main__":
    t_start = time.time()
    """ 获取名单 """
    dict_name = get_name_list()
    """ 获取工信部dict """
    file_gongxin = f'{os.path.dirname(__file__)}\\工信部清单20191031.xlsx'
    # file_gongxin = input(f'请输入《工信部清单20191031》文件名：（默认：{file_gongxin}）\n').strip() or file_gongxin
    df_gongxin = get_specify_df(file_gongxin, '10月', ['处理工号', '工单编号'])
    dict_gongxin = specify_df_frequency(df_gongxin, '处理工号').to_dict(orient='index')
    n = 0
    for i in dict_name:
        """ 《催装催修》《绿通》《抱怨》《工信部》等拿过来时少了个0 """
        if n == 0:
            print('已为部分Excel中的工号进行匹配前的适配（部分工号前部缺少0）')
        i1 = i[1:] if i.startswith('0') else i
        dict_name[i]['工信部'] = dict_gongxin.get(i1, {0:0})[0]
        
        n+=1
                
    df = pd.DataFrame.from_dict(dict_name, orient='index')
    # print(df)
    temp_excel_file = f'{os.path.dirname(__file__)}\中间表：人员工信：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    df.to_excel(excel_writer = temp_excel_file, index = True)
    print(f'已完成，保存地址{temp_excel_file}\n总耗时{time.time() - t_start}秒')
    
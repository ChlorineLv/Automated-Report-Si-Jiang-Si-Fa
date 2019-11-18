# -*- coding: UTF-8 -*-
'''
### Version: Python 3.7.4
### Date: 2019-11-14 09:53:16
### LastEditors: ChlorineLv@outlook.com
### LastEditTime: 2019-11-18 15:11:54
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


def get_zhuang_list():
    '''
    ### description: 
    ### param {type} 
    ### return: 以工号为key的各人装机次数dict
    ### example: {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    csv_file = f'{os.path.dirname(__file__)}\表1装移机工单一览表.csv'
    # csv_file = input(f'请输入《装移机工单一览表》文件名：（默认：{csv_file}）\n').strip() or csv_file
    print(f'正在处理:{csv_file}')
    # 下载下来的csv像是gbk格式
    df = pd.DataFrame(pd.read_csv(csv_file, encoding='gbk'))[['施工类型','处理人工号']]
    # 取反，取施工类型不包括拆机、其他、移拆机
    df = df[~df['施工类型'].isin(['拆机'])]
    df = df[~df['施工类型'].isin(['其他'])]
    df = df[~df['施工类型'].isin(['移拆'])]
    # print(df)
    # 以处理人工号计算频次
    df = df.groupby(by='处理人工号').size()
    # 先转为dataframe再转为dict
    return df.to_frame().to_dict(orient='index')


def get_xiu_list():
    '''
    ### description: 
    ### param {type} 
    ### return: 以工号为key的各人修障次数dict 
    ### example: {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    csv_file = f'{os.path.dirname(__file__)}\表2障碍用户申告一览表.csv'
    # csv_file = input(f'请输入《表2障碍用户申告一览表》文件名：（默认：{csv_file}）\n').strip() or csv_file
    print(f'正在处理:{csv_file}')
    # 下载下来的csv像是gbk格式
    df = pd.DataFrame(pd.read_csv(csv_file, encoding='gbk'))[['部门分局','修理员工号']]
    # 取反，取部门分局不包括合作方（甘肃万维）
    df = df[~df['部门分局'].isin(['合作方（甘肃万维）'])]
    # 取反，取部门分局不包括合作方（天讯）
    df = df[~df['部门分局'].isin(['合作方（天讯）'])]
    # print(df)
    # 以处理人工号计算频次
    df = df.groupby(by='修理员工号').size()
    # 先转为dataframe再转为dict
    return df.to_frame().to_dict(orient='index')

def get_cui_list(sheet_cui_name, column_cui_name, column_people_name):
    '''
    ### description: 
    ### param {type} 
    ### return:  以工号为key的各人催装or催装大于等于4次的单数dict
    ### example: {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    excel_file = f'{os.path.dirname(__file__)}\催装催修清单10月.xlsx'
    # excel_file = input(f'请输入《催装催修清单10月》文件名：(默认：{excel_file})\n').strip() or excel_file
    print(f'正在处理{column_cui_name}:{excel_file}')
    # 只保留列‘处理人工号’/‘修理员工号’，‘催x次数’/‘前台催x次数’
    df = pd.DataFrame(pd.read_excel(excel_file, sheet_name=sheet_cui_name))[[column_people_name, column_cui_name, '录音开始时间']]
    # print(df)
    df = df.fillna(-1)
    # 字段内容筛选‘催x次数’/‘前台催x次数’≥4次，且录音为空的
    df = df.loc[df[column_cui_name] >= 4]
    df = df.loc[df['录音开始时间'] == -1]
    # 以‘处理人工号’/‘修理员工号’计算频次
    df = df.groupby(by=column_people_name).size()
    # temp_excel_file = f'{os.path.dirname(__file__)}\{column_cui_name}{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    # df.to_excel(excel_writer = temp_excel_file, index = True)
    # 先转为dataframe再转为dict
    return df.to_frame().to_dict(orient='index')


if __name__ == "__main__":
    # 获取名单
    dict_name = get_name_list()
    # 获取装移工作量
    dict_zhuang = get_zhuang_list()
    # 获取修障工作量
    dict_xiu = get_xiu_list()
    # 获取催装大于4次的单数
    dict_cuizhuang = get_cui_list('催装', '前台催装次数', '处理人工号')
    print(dict_cuizhuang)
    # 获取催修大于4次的单数
    dict_cuixiu = get_cui_list('催修', '前台催修次数', '修理员工号')
    print(dict_cuixiu)
    n = 0
    for i in dict_name:
        dict_name[i]['装移机'] = dict_zhuang.get(i, {0:0})[0]
        dict_name[i]['修障'] = dict_xiu.get(i, {0:0})[0]
        dict_name[i]['光衰整治'] = 0
        dict_name[i]['合计'] = dict_name[i]['装移机'] + dict_name[i]['修障'] + dict_name[i]['光衰整治']
        dict_name[i]['合计（日均8，20工作日）'] = 0 if dict_name[i]['合计']/20 < 8 else 1
        if n == 0:
            # 《催装催修》拿过来时少了个0
            print('已为《催装催修》中头部缺少0的工号补齐……')
        if i.startswith('0'):
            i1 = i[1:]
        else:
            i1 = i
        dict_name[i]['催修≥4'] = dict_cuixiu.get(i1, {0:0})[0]
        dict_name[i]['催装≥4'] = dict_cuizhuang.get(i1, {0:0})[0]
        
        
        n+=1
                
    df = pd.DataFrame.from_dict(dict_name, orient='index')
    print(df)
    temp_excel_file = f'{os.path.dirname(__file__)}\中间表：人员产能表：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    df.to_excel(excel_writer = temp_excel_file, index = True)
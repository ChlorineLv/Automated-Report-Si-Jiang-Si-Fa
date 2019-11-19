# -*- coding: UTF-8 -*-
'''
:Version: Python 3.7.4
:Date: 2019-11-14 09:53:16
:LastEditors: ChlorineLv@outlook.com
:LastEditTime: 2019-11-19 17:29:57
:Description: Null
'''

import os
import pandas as pd
import time

def get_name_dict():
    '''
    :description: 以工号为key的人员名单
    :param {type} : 
    :return: {'工号1': {'姓名': '1', '手机': 1, '分公司': '1', ...}, '工号2': {'姓名': '2', '手机': 2, '分公司': '2', ...}, ......}
    '''
    excel_file = f'{os.path.dirname(__file__)}\广州电信分公司_用户管理导出(2019-11-04).xlsx'
    # excel_file = input(f'请输入《人员名单》文件名：(默认：{excel_file})\n').strip() or excel_file
    print(f'正在处理:{excel_file}')
    """ 只保留下面的列 """
    df = pd.DataFrame(pd.read_excel(excel_file))[['综调登录工号', '姓名', '手机', '分公司', '单位名称', '人员类别', '岗位类型', '在职状态']]
    """ 字段内容筛选 """
    df = df.loc[df['岗位类型'] == '外线施工岗'].loc[df['在职状态'] == '在职']
    """ 将手机号完整记录，不再是科学记数法 """
    df['手机'] = df['手机'].astype('int64')
    """ 工号列变index """
    df.set_index('综调登录工号', inplace=True)
    # print(df)
    """ 以index作为dict的key """
    return df.to_dict(orient='index')


def get_zhuang_dict():
    '''
    :description: 以工号为key的各人装机次数
    :param {type} :
    :return: {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    csv_file = f'{os.path.dirname(__file__)}\表1装移机工单一览表.csv'
    # csv_file = input(f'请输入《装移机工单一览表》文件名：（默认：{csv_file}）\n').strip() or csv_file
    print(f'正在处理:{csv_file}')
    df = pd.DataFrame(pd.read_csv(csv_file, encoding='gbk'))[['施工类型','处理人工号']]
    """ 取反，取施工类型不包括拆机、其他、移拆机 """
    df = df[~df['施工类型'].isin(['拆机'])]
    df = df[~df['施工类型'].isin(['其他'])]
    df = df[~df['施工类型'].isin(['移拆'])]
    # print(df)
    """ 以处理人工号计算频次 """
    df = df.groupby(by='处理人工号').size()
    """ 先转为dataframe再转为dict """
    return df.to_frame().to_dict(orient='index')


def get_xiu_dict():
    '''
    ### description: 以工号为key的各人修障次数
    ### param {type} 
    ### return: dict 
    ### example: {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    csv_file = f'{os.path.dirname(__file__)}\表2障碍用户申告一览表.csv'
    # csv_file = input(f'请输入《表2障碍用户申告一览表》文件名：（默认：{csv_file}）\n').strip() or csv_file
    print(f'正在处理:{csv_file}')
    df = pd.DataFrame(pd.read_csv(csv_file, encoding='gbk'))[['部门分局','修理员工号']]
    """ 取反，取部门分局不包括合作方（甘肃万维）、合作方（天讯） """
    df = df[~df['部门分局'].isin(['合作方（甘肃万维）'])]
    df = df[~df['部门分局'].isin(['合作方（天讯）'])]
    # print(df)
    """ 以处理人工号计算频次 """
    df = df.groupby(by='修理员工号').size()
    """ 先转为dataframe再转为dict """
    return df.to_frame().to_dict(orient='index')

def get_cui_dict(sheet_cui_name, column_cui_name, column_people_name):
    '''
    ### description: 以工号为key的各人催装or催装大于等于4次的工单数
    ### param {type} 
    ### return:  dict
    ### example: {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    excel_file = f'{os.path.dirname(__file__)}\催装催修清单10月.xlsx'
    # excel_file = input(f'请输入《催装催修清单10月》文件名：(默认：{excel_file})\n').strip() or excel_file
    print(f'正在处理{column_cui_name}:{excel_file}')
    """ 只保留列‘处理人工号’/‘修理员工号’，‘催x次数’/‘前台催x次数’ """
    df = pd.DataFrame(pd.read_excel(excel_file, sheet_name=sheet_cui_name))[[column_people_name, column_cui_name, '录音开始时间']]
    # print(df)
    df = df.fillna(-1)
    """ 字段内容筛选‘催x次数’/‘前台催x次数’≥4次，且录音为空的 """
    df = df.loc[df[column_cui_name] >= 4]
    df = df.loc[df['录音开始时间'] == -1]
    """ 以‘处理人工号’/‘修理员工号’计算频次 """
    df = df.groupby(by=column_people_name).size()
    """ 先转为dataframe再转为dict """
    return df.to_frame().to_dict(orient='index')


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


def specify_df_baoyuan(df_get, c_name):
    '''
    :description: 统计抱怨/绿通dataframe中各人抱怨量
    :param df_get {dataframe} :
    :param c_name {str} : 工号字段具体名称
    :return: dict {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    # 以处理人工号计算频次
    df = df_get.groupby(by=c_name).size()
    return df.to_frame().to_dict(orient='index')


if __name__ == "__main__":
    t_start = time.time()
    """ 获取名单 """
    dict_name = get_name_dict()
    """ 获取装移工作量 """
    dict_zhuang = get_zhuang_dict()
    """ 获取修障工作量 """
    dict_xiu = get_xiu_dict()
    """ 获取催装大于4次的单数 """
    dict_cuizhuang = get_cui_dict('催装', '前台催装次数', '处理人工号')
    # print(dict_cuizhuang)
    """ 获取催修大于4次的单数 """
    dict_cuixiu = get_cui_dict('催修', '前台催修次数', '修理员工号')
    # print(dict_cuixiu)
    """ 获取《抱怨》的dataframe """
    file_baoyuan = f'{os.path.dirname(__file__)}\抱怨清单统计2019年10月.xlsx'
    # file_baoyuan = input(f'请输入《抱怨清单统计2019年10月》文件名：（默认：{file_baoyuan}）\n').strip() or file_baoyuan
    df_baoyuan = get_specify_df(file_baoyuan, '清单', ['工号', '服务类别'])
    """ 获取《绿通》的dataframe """
    file_lvtong = f'{os.path.dirname(__file__)}\绿通单清单2019年10月.xlsx'
    # file_lvtong = input(f'请输入《抱怨清单统计2019年10月》文件名：（默认：{file_lvtong}）\n').strip() or file_lvtong
    df_lvtong = get_specify_df(file_lvtong, '1', ['处理人工号','服务类型'])
    """ 获取各人《抱怨》总量和《绿通》总量 """
    dict_baoyuan_total = specify_df_baoyuan(df_baoyuan, '工号')
    dict_lvtong_total = specify_df_baoyuan(df_lvtong, '处理人工号')
    n = 0
    for i in dict_name:
        dict_name[i]['装移机'] = dict_zhuang.get(i, {0:0})[0]
        dict_name[i]['修障'] = dict_xiu.get(i, {0:0})[0]
        dict_name[i]['光衰整治'] = 0
        dict_name[i]['合计'] = dict_name[i]['装移机'] + dict_name[i]['修障'] + dict_name[i]['光衰整治']
        dict_name[i]['合计（日均8，20工作日）'] = 0 if dict_name[i]['合计']/20 < 8 else 1
        # 《催装催修》拿过来时少了个0
        if n == 0:
            print('已为《催装催修》中头部缺少0的工号进行匹配前的适配……')
        i1 = i[1:] if i.startswith('0') else i
        dict_name[i]['催修≥4'] = dict_cuixiu.get(i1, {0:0})[0]
        dict_name[i]['催装≥4'] = dict_cuizhuang.get(i1, {0:0})[0]
        dict_name[i]['抱怨'] = dict_baoyuan_total.get(i1, {0:0})[0] + dict_lvtong_total.get(i1, {0:0})[0]
        
        n+=1
                
    df = pd.DataFrame.from_dict(dict_name, orient='index')
    print(df)
    temp_excel_file = f'{os.path.dirname(__file__)}\中间表：人员产能表：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    df.to_excel(excel_writer = temp_excel_file, index = True)
    print(f'已完成，保存地址{temp_excel_file}\n总耗时{time.time() - t_start}秒')
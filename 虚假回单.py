# -*- coding: UTF-8 -*-
'''
:Version: Python 3.7.4
:Date: 2019-11-25 17:18:56
:LastEditors: ChlorineLv@outlook.com
:LastEditTime: 2019-12-18 10:36:51
:Description: 虚假回单
'''
import os
import pandas as pd
import time

def get_name_from(excel_file):
    '''
    :description: 从已有《四奖四罚》中获得人员名单
    :param excel_name {str} :
    :param sheet_name {str} :
    :param column_list {str} :取得的列名list
    :return: dict
    '''
    print(f'正在处理:{excel_file}\n')
    # 只保留下面的列
    df = pd.DataFrame(pd.read_excel(excel_file, sheet_name='人员产能情况'))[['代维工号', '姓名', '工号', '手机号', '所属公司', '区域', '装维班', '人员类别', '人员属性', '岗位']]
    df['代维工号'] = df['代维工号'].map(lambda x: str(x)[1:] if str(x).startswith('0') else str(x))
    df.set_index('代维工号', inplace=True)
    return df.to_dict(orient='index')


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
    df['综调登录工号'] = df['综调登录工号'].map(lambda x: str(x)[1:] if str(x).startswith('0') else str(x))
    # 工号列变index
    df.set_index('综调登录工号', inplace=True)
    # print(df)
    # 以index作为dict的key
    return df.to_dict(orient='index')


def get_specified_df(excel_file, excel_sheet, column_list=0):
    '''
    :description: 返回保留某几列的dataframe
    :param excel_file {str} : excel文件所在位置
    :param excel_sheet {str} : sheet名称
    :param column_list {list} : 工号等列名
    :return: dataframe
    '''
    print(f'正在处理:{excel_file}')
    if column_list != 0:
        df = pd.DataFrame(pd.read_excel(excel_file, sheet_name=excel_sheet, dtype={column_list[0]: str}))[column_list]
    else:
        df = pd.DataFrame(pd.read_excel(excel_file, sheet_name=excel_sheet, dtype={column_list[0]: str}))
    df[column_list[0]] = df[column_list[0]].map(lambda x: str(x)[1:] if str(x).startswith('0') else str(x))
    # print(df)
    return df


def specify_df_frequency(df_get, c_name, c_judge=0):
    '''
    :description: 返回df中各项含特定字段的频次
    :param df_get {dataframe} : dataframe
    :param c_name {str} : 需要保留的字段
    :param c_judge {dict} : 用于判断的列:值对，如{'服务3':'超时未修复故障'}
    :return: dict {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    """ 将工号等转为string格式 """
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
    return df.to_dict(orient='index')


def specify_df_frequency_or(df_get, c_name, c_judge=[]):
    '''
    :description: 返回df中各项含任一特定字段的频次
    :param df_get {dataframe} : dataframe
    :param c_name {str} : 需要保留的字段
    :param c_judge {list} : 用于判断的列:值对的集合，如[{'服务3':'超时未修复故障'}, {'服务3':'未按预约时间上门'}]
    :return: dict {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    """ 将工号等转为string格式 """
    df = pd.DataFrame()
    df_get[c_name] = df_get[c_name].apply(str)
    if len(c_judge) == 0:
        df = df_get.groupby(by=c_name).count()
    elif len(c_judge[0]) == 0:
        df = df_get.groupby(by=c_name).count()
    else:
        for i in range(len(c_judge)):
            for (k,v) in c_judge[i].items():
                key = k
                value = c_judge[i][k]
            df = df.append(df_get.loc[df_get[key]==value][[c_name, key]])
    df = df.groupby(by=c_name).count()
    df.columns = [0]
    # temp_excel_file = f'{os.path.dirname(__file__)}\中间表：{c_name}{value}：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    # df.to_excel(excel_writer = temp_excel_file, index = True)
    return df.to_dict(orient='index')


def specify_df_frequency_not_in_or(df_get, c_name, c_judge=[]):
    '''
    :description: 返回df中各项不含特定字段，不同内容的频次
    :param df_get {dataframe} : dataframe
    :param c_name {str} : 需要保留的字段
    :param c_judge {list} : 用于判断不包含的列:值对的集合，如[{'服务3':'1'}, {'服务3':'2'}]
    :return: dict {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    """ 将工号等转为string格式 """
    df = pd.DataFrame()
    df_get[c_name] = df_get[c_name].apply(str)
    if len(c_judge) == 0:
        df = df_get.groupby(by=c_name).count()
    elif len(c_judge[0]) == 0:
        df = df_get.groupby(by=c_name).count()
    else:
        for i in range(len(c_judge)):
            for (k,v) in c_judge[i].items():
                key = k
                value = c_judge[i][k]
            df_get = df_get[~df_get[key].isin([value])]
        """ 去除NaN项 """
        df_get.fillna(-1)
        df_get = df_get[~df_get[key].isin([-1])]
    df = df_get[[c_name, key]]
    df = df.groupby(by=c_name).count()
    df.columns = [0]
    # temp_excel_file = f'{os.path.dirname(__file__)}\中间表：{c_name}{value}：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    # df.to_excel(excel_writer = temp_excel_file, index = True)
    return df.to_dict(orient='index')


def specify_df_in(df_get, c_judge):
    '''
    :description: 筛选包括的行，并返回dataframe
    :param c_judge {dict} : 例如筛选服务3为1的行，则{'服务3':'1'}
    :return: dataframe
    '''
    for (k,v) in c_judge.items():
        key = k
        value = c_judge[k]
    df_get = df_get[df_get[key]==value]
    return df_get


def specify_df_not_include(df_get, c_judge):
    '''
    :description: 筛选去掉含有某个字符的行（如含1的行，有可能11，111，1111）
    :param c_judge {dict} : 例如去掉服务3含有1的行，则{'服务3':'1'}
    :return: dataframe
    '''
    for (k,v) in c_judge.items():
        key = k
        value = c_judge[k]
    df_get.loc[~df_get[key].str.contains(value)]
    return df_get


if __name__ == "__main__":
    t_start = time.time()
    """ 获取名单 """
    temp_exist_file = '装维人员服务奖惩通报（8月）.xlsx'
    dict_name = get_name_from(temp_exist_file)
    # dict_name = get_name_list()
    """ 获取ivr故障虚假回单dict """
    file_xujia_ivr_guzhang = f'{os.path.dirname(__file__)}\\（IVR）故障虚假回单清单8月.xlsx'
    # file_xujia_ivr_guzhang = input(f'请输入《（IVR）故障虚假回单清单2019年10月》文件名：（默认：{file_xujia_ivr_guzhang}）\n').strip() or file_xujia_ivr_guzhang
    sheet_xujia_ivr_guzhang = '工单明细'
    # sheet_xujia_ivr_guzhang = input(f'请输入sheet名：（默认{sheet_xujia_ivr_guzhang})\n').strip() or sheet_xujia_ivr_guzhang
    df_xujia_ivr_guzhang = get_specified_df(file_xujia_ivr_guzhang, sheet_xujia_ivr_guzhang, ['查修员工号', 'B、请问您的故障修复了吗？'])
    dict_xujia_ivr_guzhang = specify_df_frequency(df_xujia_ivr_guzhang, '查修员工号', {'B、请问您的故障修复了吗？': '未修复，请按3'})
    """ 获取ivr装机虚假回单dict """
    file_xujia_ivr_zhuangji = f'{os.path.dirname(__file__)}\\（IVR）装机虚假回单清单8月.xlsx'
    # file_xujia_ivr_zhuangji = input(f'请输入《（IVR）装机虚假回单清单2019年10月》文件名：（默认：{file_xujia_ivr_zhuangji}）\n').strip() or file_xujia_ivr_zhuangji
    sheet_xujia_ivr_zhuangji = '工单明细'
    # sheet_xujia_ivr_zhuangji = input(f'请输入sheet名：（默认{sheet_xujia_ivr_zhuangji})\n').strip() or sheet_xujia_ivr_zhuangji
    df_xujia_ivr_zhuangji = get_specified_df(file_xujia_ivr_zhuangji, sheet_xujia_ivr_zhuangji, ['装维人员工号', 'B、请问您的电信业务能正常使用吗？'])
    dict_xujia_ivr_zhuangji = specify_df_frequency_not_in_or(df_xujia_ivr_zhuangji, '装维人员工号',[{'B、请问您的电信业务能正常使用吗？': '能正常使用，请按2'}])
    """ 获取人工故障虚假回单dict """
    file_xujia_rengong_guzhang = f'{os.path.dirname(__file__)}\\（人工）故障虚假回单清单8月.xlsx'
    # file_xujia_rengong_guzhang = input(f'请输入《（人工）故障虚假回单清单2019年10月》文件名：（默认：{file_xujia_rengong_guzhang}）\n').strip() or file_xujia_rengong_guzhang
    sheet_xujia_rengong_guzhang = 'Sheet1'
    # sheet_xujia_rengong_guzhang = input(f'请输入sheet名：（默认{sheet_xujia_rengong_guzhang})\n').strip() or sheet_xujia_rengong_guzhang
    df_xujia_rengong_guzhang = get_specified_df(file_xujia_rengong_guzhang, sheet_xujia_rengong_guzhang, ['查修员工号', 'D、请问维修人员有联系您处理过吗？'])
    dict_xujia_rengong_guzhang = specify_df_frequency(df_xujia_rengong_guzhang, '查修员工号', {'D、请问维修人员有联系您处理过吗？': '一直无人联系修障'})
    """ 获取人工装机虚假回单dict """
    file_xujia_rengong_zhuangji = f'{os.path.dirname(__file__)}\\（人工）装机虚假回单清单8月.xlsx'
    # file_xujia_rengong_zhuangji = input(f'请输入《（人工）故障虚假回单清单2019年10月》文件名：（默认：{file_xujia_rengong_zhuangji}）\n').strip() or file_xujia_rengong_zhuangji
    sheet_xujia_rengong_zhuangji = 'Sheet1'
    # sheet_xujia_rengong_zhuangji = input(f'请输入sheet名：（默认{sheet_xujia_rengong_zhuangji})\n').strip() or sheet_xujia_rengong_zhuangji
    column_xujia_rengong_zhuangji = ['装维人员工号', 'B、请问您的电信业务能正常使用吗？', 'E、请问是哪种情况不能使用？', 'F、请问是什么原因没当场安装好呢？']
    # df_xujia_rengong_zhuangji = get_specified_df(file_xujia_rengong_zhuangji, sheet_xujia_rengong_zhuangji, ['装维人员工号', 'B、请问您的电信业务能正常使用吗？'])
    df_xujia_rengong_zhuangji = get_specified_df(file_xujia_rengong_zhuangji, sheet_xujia_rengong_zhuangji, column_xujia_rengong_zhuangji)
    df_xujia_rengong_zhuangji = specify_df_in(df_xujia_rengong_zhuangji, {'E、请问是哪种情况不能使用？':'没当场装好，至今不能正常使用'})
    df_xujia_rengong_zhuangji = specify_df_not_include(df_xujia_rengong_zhuangji, {'B、请问您的电信业务能正常使用吗？': '能正常使用'})
    dict_xujia_rengong_zhuangji = specify_df_frequency_not_in_or(df_xujia_rengong_zhuangji, '装维人员工号',[{'F、请问是什么原因没当场安装好呢？': '我不需要安装电视机顶盒'}])

    n = 0
    for i in dict_name:

        dict_name[i]['ivr故障虚假回单'] = dict_xujia_ivr_guzhang.get(i, {0:0})[0] 
        dict_name[i]['ivr装机虚假回单'] = dict_xujia_ivr_zhuangji.get(i, {0:0})[0]
        dict_name[i]['人工故障虚假回单'] = dict_xujia_rengong_guzhang.get(i, {0:0})[0]
        dict_name[i]['人工装机虚假回单'] = dict_xujia_rengong_zhuangji.get(i, {0:0})[0]
        
        n+=1
                
    df = pd.DataFrame.from_dict(dict_name, orient='index')
    # print(df)
    temp_excel_file = f'{os.path.dirname(__file__)}\中间表：：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    df.to_excel(excel_writer = temp_excel_file, index = True)
    print(f'已完成，保存地址{temp_excel_file}\n总耗时{time.time() - t_start}秒')
    
# -*- coding: UTF-8 -*-
'''
:Version: Python 3.7.4
:Date: 2019-11-28 17:26:05
:LastEditors: ChlorineLv@outlook.com
:LastEditTime: 2019-12-17 14:26:52
:Description: 满意度
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
    df['综调登录工号'] = df['综调登录工号'].map(lambda x: str(x)[1:] if str(x).startswith('0') else str(x))
    # 工号列变index
    df.set_index('综调登录工号', inplace=True)
    # print(df)
    # 以index作为dict的key
    return df.to_dict(orient='index')


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
    print("\n",df.loc[df[column_list[0]]=='01038315'],"\n")
    print('\n', df.iloc[7987], df.iloc[7988], df.iloc[7989], df.iloc[7990], df.iloc[7991], df.iloc[7992])
    df[column_list[0]] = df[column_list[0]].map(lambda x: str(x)[1:] if str(x).startswith('0') else str(x))
    print(df)
    return df


def specify_df_manyi(df_get, column_name, column_judge):
    '''
    :description: 满意度统计
    :param df_get {dataframe} : dataframe
    :param column_name {str} : 需要保留的字段
    :param column_judge {list} : 用于判断不包含的列:值对的集合，如 [{'满意度':'非常满意'}, {'满意度':'满意'}, {'满意度':'不满意'}]
    :return: [{'工号': {'非常满意': 4.0, '满意': 1.0, '不满意': 0}, ...]
    '''
    df_get[column_name] = df_get[column_name].apply(str)
    df = pd.DataFrame()
    for (k,v) in column_judge[i].items():
        key = k
        value = column_judge[i][k]
    df_get = df_get.loc[df_get[key]==value].groupby(column_name).count()
    # for i in range(len(column_judge)):
    #     for (k,v) in column_judge[i].items():
    #         key = k
    #         value = column_judge[i][k]
    #     df_temp = df_get.loc[df_get[key]==value].groupby(column_name).count()
    #     # print(k,value, df_temp)
    #     df_temp.columns=[value]
    #     df = df.join(df_temp, how='outer')
    #     df.fillna(0, inplace=True)
    #     df[value] = df[value].apply(int)
    # print(df)
    return df.to_dict(orient='index')


def specify_df_frequency(df_get, column_name, column_judge=0):
    '''
    :description: 返回df中各项含特定字段的频次
    :param df_get {dataframe} : dataframe
    :param column_name {str} : 需要保留的字段
    :param column_judge {dict} : 用于判断的列:值对，如{'服务3':'超时未修复故障'}
    :return: dict {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    """ 将工号等转为string格式 """
    df_get[column_name] = df_get[column_name].apply(str)
    if column_judge == 0:
        df = df_get.groupby(by=column_name).count()
    else:
        for (k,v) in column_judge.items():
            key = k
            value = column_judge[k]
        # testid = '1038315'
        # print(df_get.loc[df_get[column_name]==testid])
        
        df = df_get.loc[df_get[key]==value][[column_name, key]].groupby(by=column_name).count()
    df.columns = [0]
    # temp_excel_file = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\中间表：{column_name}{value}：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    # df.to_excel(excel_writer = temp_excel_file, index = True)
    return df.to_dict(orient='index')


if __name__ == "__main__":
    t_start = time.time()
    pd.set_option('display.max_rows', 500)
    # month = input('请输入当前月份：（如10、11）\n').strip()
    month = '8'
    """ 获取名单 """
    # dict_name = get_name_list()
    # temp_exist_file = input(f'\n请输入已有四奖四罚文件名，以获取已有人员名单：\n').strip()
    temp_exist_file = '装维人员服务奖惩通报（8月）.xlsx'
    dict_name = get_name_from(temp_exist_file)
    """ 修障满意度 """
    file_manyi_xiuzhang = f'{os.path.dirname(__file__)}\\修障服务测评清单（含IVR&人工）{month}月.xlsx'
    # file_manyi_xiuzhang = input(f'请输入《修障服务测评清单（含IVR&人工）2019年10月》文件名：（默认：{file_manyi_xiuzhang}）\n').strip() or file_manyi_xiuzhang
    sheet_manyi_xiuzhang = 'Sheet1'
    # sheet_manyi_xiuzhang = input(f'请输入sheet名：（默认{sheet_manyi_xiuzhang})\n').strip() or sheet_manyi_xiuzhang
    df_manyi_xiuzhang = get_specified_df(file_manyi_xiuzhang, sheet_manyi_xiuzhang, ['查修员工号', '满意度','工单号'])
    # dict_manyi_xiuzhang = specify_df_manyi(df_manyi_xiuzhang, '操作工号', [{'满意度':'非常满意'}, {'满意度':'满意'}, {'满意度':'不满意'}])
    dict_manyi_xiuzhang_feichang = specify_df_frequency(df_manyi_xiuzhang, '查修员工号', {'满意度':'非常满意'})
    dict_manyi_xiuzhang_manyi = specify_df_frequency(df_manyi_xiuzhang, '查修员工号', {'满意度':'满意'})
    dict_manyi_xiuzhang_buman = specify_df_frequency(df_manyi_xiuzhang, '查修员工号', {'满意度':'不满意'})    
    # print(dict_manyi_xiuzhang)
    """ 装机满意度 """
    file_manyi_zhuangji = f'{os.path.dirname(__file__)}\\装机服务测评清单（含IVR&人工）{month}月.xlsx'
    # file_manyi_zhuangji = input(f'请输入《装机服务测评清单（含IVR&人工）2019年10月》文件名：（默认：{file_manyi_zhuangji}）\n').strip() or file_manyi_zhuangji
    sheet_manyi_zhuangji = 'Sheet1'
    # sheet_manyi_zhuangji = input(f'请输入sheet名：（默认{sheet_manyi_zhuangji})\n').strip() or sheet_manyi_zhuangji
    df_manyi_zhuangji = get_specified_df(file_manyi_zhuangji, sheet_manyi_zhuangji, ['装维人员工号', '满意度','工单号'])
    # dict_manyi_zhuangji = specify_df_manyi(df_manyi_zhuangji, '装维人员工号', [{'满意度':'非常满意'}, {'满意度':'满意'}, {'满意度':'不满意'}])
    dict_manyi_zhuangji_feichang = specify_df_frequency(df_manyi_zhuangji, '装维人员工号', {'满意度':'非常满意'})
    dict_manyi_zhuangji_manyi = specify_df_frequency(df_manyi_zhuangji, '装维人员工号', {'满意度':'满意'})
    dict_manyi_zhuangji_bumanyi = specify_df_frequency(df_manyi_zhuangji, '装维人员工号', {'满意度':'不满意'})

    
    n = 0
    for i in dict_name:
        dict_name[i]['修非常满意'] = dict_manyi_xiuzhang_feichang.get(i, {0:0})[0]
        dict_name[i]['修满意'] = dict_manyi_xiuzhang_manyi.get(i, {0:0})[0]
        dict_name[i]['修不满意'] = dict_manyi_xiuzhang_buman.get(i, {0:0})[0]
        dict_name[i]['装非常满意'] = dict_manyi_zhuangji_feichang.get(i, {0:0})[0]
        dict_name[i]['装满意'] = dict_manyi_zhuangji_manyi.get(i, {0:0})[0]
        dict_name[i]['装不满意'] = dict_manyi_zhuangji_bumanyi.get(i, {0:0})[0]
        n+=1
                
    df = pd.DataFrame.from_dict(dict_name, orient='index')
    # print(df)
    temp_excel_file = f'{os.path.dirname(__file__)}\中间表：满意度：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    df.to_excel(excel_writer = temp_excel_file, index = True)
    print(f'已完成，保存地址{temp_excel_file}\n总耗时{time.time() - t_start}秒')
# -*- coding: UTF-8 -*-
'''
:Version: Python 3.7.4
:Date: 2019-11-14 09:53:16
:LastEditors: ChlorineLv@outlook.com
:LastEditTime: 2019-12-30 15:21:03
:Description: Null
'''

import os
import signal
import sys
import time

import pandas as pd


def signal_handler(signal, frame):
    print("键盘输入了: Ctrl C")
    if signal==2 and input("输入任意字符以退出脚本（回车除外）：")!='':
        sys.exit(0)


def get_name_list(excel_file):
    '''
    ### description: 
    ### param {type} 
    ### return: 以工号为key的人员名单dict
    ### example: {'工号1': {'姓名': '1', '手机': 1, '分公司': '1', ...}, '工号2': {'姓名': '2', '手机': 2, '分公司': '2', ...}, ......}
    '''
    
    print(f'正在处理:{excel_file}\n')
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


def get_zhuang_dict(excel_file):
    '''
    :description: 以工号为key的各人装机次数
    :param {type} :
    :return: {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    print(f'正在处理“装移机工单量”: {excel_file}\n')
    df = pd.DataFrame(pd.read_excel(excel_file, sheet_name='匹配清单'))[['施工类型','处理人工号', '归档模式']]
    df['处理人工号'] = df['处理人工号'].map(lambda x: str(x)[1:] if str(x).startswith('0') else str(x))
    """ 取反，取施工类型不包括拆机、其他、移拆机 """
    df_temp = pd.DataFrame()
    df_temp = df_temp.append(df[df['施工类型']=='改信息'])
    df_temp = df_temp.append(df[df['施工类型']=='同楼移机'])
    df_temp = df_temp.append(df[df['施工类型']=='新装'])
    df_temp = df_temp.append(df[df['施工类型']=='移机'])
    df_temp = df_temp.append(df[df['施工类型']=='移装'])
    df_temp = df_temp.loc[df_temp['归档模式']=='正常归档']

    df = df_temp
    # print(df)
    """ 以处理人工号计算频次 """
    df = df.groupby(by='处理人工号').size().to_frame()
    """ 先在前面转为dataframe再转为dict """
    return df.to_dict(orient='index')


def get_xiu_dict(excel_file):
    '''
    ### description: 以工号为key的各人修障次数
    ### param {type} 
    ### return: dict 
    ### example: {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    print(f'正在处理“修障工单量”: {excel_file}\n')
    df = pd.DataFrame(pd.read_excel(excel_file, sheet_name='匹配清单'))[['部门分局','修理员工号','专业名称']]
    df['修理员工号'] = df['修理员工号'].map(lambda x: str(x)[1:] if str(x).startswith('0') else str(x))
    """ 取反，剔除专业名称：现场检修 """
    df = df[~df['专业名称'].isin(['现场检修'])]
    df = df[~df['专业名称'].isin(['预检预修'])]
    df = df[['部门分局','修理员工号']]
    """ 取反，取部门分局不包括合作方（甘肃万维）、合作方（天讯） """
    df = df[~df['部门分局'].isin(['合作方（甘肃万维）'])]
    df = df[~df['部门分局'].isin(['合作方（天讯）'])]
    # print(df)
    """ 以处理人工号计算频次 """
    df = df.groupby(by='修理员工号').size().to_frame()
    """ 先在前面转为dataframe再转为dict """
    return df.to_dict(orient='index')


def get_specified_df(excel_file, excel_sheet, column_list=0):
    '''
    :description: 返回保留某几列的dataframe
    :param excel_file {str} : excel文件所在位置
    :param excel_sheet {str} : sheet名称
    :param column_list {list} : 工号等列名
    :return: dataframe
    '''
    print(f'正在处理:{excel_file}\n')
    if column_list != 0:
        df = pd.DataFrame(pd.read_excel(excel_file, sheet_name=excel_sheet, dtype={column_list[0]: str}))[column_list]
    else:
        df = pd.DataFrame(pd.read_excel(excel_file, sheet_name=excel_sheet, dtype={column_list[0]: str}))
    df[column_list[0]] = df[column_list[0]].map(lambda x: str(x)[1:] if str(x).startswith('0') else str(x))
    # print(df)
    return df


def specify_df_cui(df, column_name, column_judge):
    '''
    :description: 筛选‘催x次数’/‘前台催x次数’≥4次，且录音为空的
    :param df_get {dataframe} : dataframe
    :param column_name {string} : 工号
    :param column_judge {string} : "催装"、"催修"
    :return: dict
    '''
    df = df.fillna(0)
    """ 字段内容筛选‘催x次数’/‘前台催x次数’≥4次，且录音为空的 """
    df = df.loc[df[column_judge] >= 4]
    df = df.loc[df['录音开始时间'] == 0]
    """ 以‘处理人工号’/‘修理员工号’计算频次 """
    df = df.groupby(by=column_name).size().to_frame()
    # print(column_name, df)
    """ 先在前面转为dataframe再转为dict """
    return df.to_dict(orient='index')


def specify_df_baoyuan(df_get, column_name):
    '''
    :description: 统计抱怨/绿通dataframe中各人抱怨量
    :param df_get {dataframe} :
    :param column_name {str} : 工号字段具体名称
    :return: dict {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    # 以处理人工号计算频次
    df = df_get.groupby(by=column_name).size().to_frame()
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
        df = df_get.loc[df_get[key]==value][[column_name, key]].groupby(by=column_name).count()
    df.columns = [0]
    # temp_excel_file = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\中间表：{column_name}{value}：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    # df.to_excel(excel_writer = temp_excel_file, index = True)
    return df.to_dict(orient='index')


def specify_df_in_or(df_get, column_judge=[]):
    '''
    :description: 返回df中各项含任一特定字段
    :param df_get {dataframe} : dataframe
    :param column_judge {list} : 用于判断的列:值对的集合，如[{'服务3':'超时未修复故障'}, {'服务3':'未按预约时间上门'}]
    :return: dataframe
    '''
    df = pd.DataFrame()
    for i in range(len(column_judge)):
        for (k,v) in column_judge[i].items():
            key = k
            value = column_judge[i][k]
        df = df.append(df_get.loc[df_get[key]==value])
    # temp_excel_file = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\中间表：{column_name}{value}：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    # df.to_excel(excel_writer = temp_excel_file, index = True)
    return df


def specify_df_frequency_or(df_get, column_name, column_judge=[]):
    '''
    :description: 返回df中各项含任一特定字段的频次
    :param df_get {dataframe} : dataframe
    :param column_name {str} : 需要保留的字段
    :param column_judge {list} : 用于判断的列:值对的集合，如[{'服务3':'超时未修复故障'}, {'服务3':'未按预约时间上门'}]
    :return: dict {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    """ 将工号等转为string格式 """
    df = pd.DataFrame()
    df_get[column_name] = df_get[column_name].apply(str)
    if len(column_judge) == 0:
        df = df_get.groupby(by=column_name).count()
    elif len(column_judge[0]) == 0:
        df = df_get.groupby(by=column_name).count()
    else:
        for i in range(len(column_judge)):
            for (k,v) in column_judge[i].items():
                key = k
                value = column_judge[i][k]
            df = df.append(df_get.loc[df_get[key]==value][[column_name, key]])
    df = df.groupby(by=column_name).count()
    df.columns = [0]
    # temp_excel_file = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\中间表：{column_name}{value}：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
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


def specify_df_frequency_not_in_or(df_get, column_name, column_judge=[]):
    '''
    :description: 返回df中各项不含特定字段，不同内容的频次
    :param df_get {dataframe} : dataframe
    :param column_name {str} : 需要保留的字段
    :param column_judge {list} : 用于判断不包含的列:值对的集合，如[{'服务3':'1'}, {'服务3':'2'}]
    :return: dict {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    """ 将工号等转为string格式 """
    df = pd.DataFrame()
    df_get[column_name] = df_get[column_name].apply(str)
    if len(column_judge) == 0:
        df = df_get.groupby(by=column_name).count()
    elif len(column_judge[0]) == 0:
        df = df_get.groupby(by=column_name).count()
    else:
        for i in range(len(column_judge)):
            for (k,v) in column_judge[i].items():
                key = k
                value = column_judge[i][k]
            df_get = df_get[~df_get[key].isin([value])]
        """ 去除NaN项 """
        df_get.fillna(-1)
        df_get = df_get[~df_get[key].isin([-1])]
    df = df_get[[column_name, key]]
    df = df.groupby(by=column_name).count()
    df.columns = [0]
    # temp_excel_file = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\中间表：{column_name}{value}：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    # df.to_excel(excel_writer = temp_excel_file, index = True)
    return df.to_dict(orient='index')


# def specify_df_manyi(df_get, column_name, column_judge):
#     '''
#     :description: 满意度统计
#     :param df_get {dataframe} : dataframe
#     :param column_name {str} : 需要保留的字段
#     :param column_judge {list} : 用于判断不包含的列:值对的集合，如 [{'满意度':'非常满意'}, {'满意度':'满意'}, {'满意度':'不满意'}]
#     :return: [{'工号': {'非常满意': 4.0, '满意': 1.0, '不满意': 0}, ...]
#     '''
#     df_get[column_name] = df_get[column_name].apply(str)
#     df = pd.DataFrame()
#     for i in range(len(column_judge)):
#         for (k,v) in column_judge[i].items():
#             key = k
#             value = column_judge[i][k]
#         df_temp = df_get.loc[df_get[key]==value].groupby(column_name).count()
#         # print(k,value, df_temp)
#         df_temp.columns=[value]
#         df = df.join(df_temp, how='outer')
#         df.fillna(0, inplace=True)
#         df[value] = df[value].apply(int)
#     return df.to_dict(orient='index')


if __name__ == "__main__":
    
    signal.signal(signal.SIGINT,signal_handler)
    t_start = time.time()
    try:
        print(f'地址脚本所在地址{os.path.dirname(os.path.realpath(sys.argv[0]))}')
        month = input('请输入当前月份：（如10、11）\n').strip()
        option = input('是否需要手动修改文件名（y/n)\n').strip()
        # if option == '111':
        #     temp_option = option
        #     temp_exist_file = input(f'\n请输入已有四奖四罚文件名，以获取已有人员名单：\n').strip()
        #     option = input('是否需要手动修改文件名（y/n)\n').strip()
        # else:
        #     temp_option = 0
        boolean = (option=='n' or option=='N')
        """ 获取名单 """
        print("\n************  处理（1/15）：人员名单  ************")
        file_name = f'人员名单（{month}月）.xlsx'
        file_name = file_name if boolean==True else (input(f'\n请输入《人员名单》文件名：(默认：{file_name})\n').strip() or file_name)
        dict_name = get_name_from(file_name)
        # dict_name = get_name_from()
        """ 获取装移工作量 """
        print("\n************  处理（2/15）：装移机工作量  ************")
        # file_zhuangji = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\表1装移机工单一览表.csv'
        file_zhuangji = f'表1装移机工单一览表.xlsx'
        file_zhuangji = file_zhuangji if boolean==True else (input(f'\n请输入《表1装移机工单一览表》文件名：（默认：{file_zhuangji}）\n').strip() or file_zhuangji)
        dict_zhuang = get_zhuang_dict(file_zhuangji)
        """ 获取修障工作量 """
        print("\n************  处理（3/15）：修障工作量  ************")
        # file_xiuzhang = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\表2障碍用户申告一览表.csv'
        file_xiuzhang = f'表2障碍用户申告一览表.xlsx'
        file_xiuzhang = file_xiuzhang if boolean==True else (input(f'\n请输入《表2障碍用户申告一览表》文件名：（默认：{file_xiuzhang}）\n').strip() or file_xiuzhang)
        dict_xiu = get_xiu_dict(file_xiuzhang)
        """ 获取催装大于4次的单数 """
        print("\n************  处理（4/15）：催装≥4  ************")
        # file_cuizhuang = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\催装催修清单10月.xlsx'
        file_cuizhuang = f'催装催修清单{month}月.xlsx'
        file_cuizhuang = file_cuizhuang if boolean==True else (input(f'\n请输入《催装催修清单{month}月》文件名：(默认：{file_cuizhuang})\n').strip() or file_cuizhuang)
        sheet_cuizhuang = '催装'
        df_cuizhuang = get_specified_df(file_cuizhuang, sheet_cuizhuang, ['处理人工号', '前台催装次数', '录音开始时间'])
        dict_cuizhuang = specify_df_cui(df_cuizhuang, '处理人工号', '前台催装次数')
        """ 获取催修大于4次的单数 """
        print("\n************  处理（5/15）：催修≥4  ************")
        # file_cuixiu = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\催装催修清单10月.xlsx'
        file_cuixiu = f'催装催修清单{month}月.xlsx'
        file_cuixiu = file_cuixiu if boolean==True else (input(f'\n请输入《催装催修清单{month}月》文件名：(默认：{file_cuixiu})\n').strip() or file_cuixiu)
        sheet_cuixiu = '催修'
        df_cuixiu = get_specified_df(file_cuixiu, sheet_cuixiu, ['修理员工号', '前台催修次数', '录音开始时间'])
        dict_cuixiu = specify_df_cui(df_cuixiu, '修理员工号', '前台催修次数')
        """ 获取《抱怨》的dataframe """
        print("\n************  处理（6/15）：《抱怨单》  ************")
        # file_baoyuan = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\抱怨清单统计{month}月.xlsx'
        file_baoyuan = f'抱怨清单统计{month}月.xlsx'
        file_baoyuan = file_baoyuan if boolean==True else (input(f'\n请输入《抱怨清单统计{month}月》文件名：（默认：{file_baoyuan}）\n').strip() or file_baoyuan)
        sheet_baoyuan = '匹配清单'
        df_baoyuan = get_specified_df(file_baoyuan, sheet_baoyuan, ['工号', '服务2', '服务3'])
        """ 获取《绿通》的dataframe """
        print("\n************  处理（7/15）：《绿通单》  ************")
        # file_lvtong = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\绿通单清单10月.xlsx'
        file_lvtong = f'绿通单清单{month}月.xlsx'
        file_lvtong = file_lvtong if boolean==True else (input(f'\n请输入《抱怨清单统计{month}月》文件名：（默认：{file_lvtong}）\n').strip() or file_lvtong)
        sheet_lvtong = '匹配清单'
        df_lvtong = get_specified_df(file_lvtong, sheet_lvtong, ['工号','服务类型'])
        """ 获取各人《抱怨》总量和《绿通》总量 """
        dict_baoyuan_total = specify_df_baoyuan(df_baoyuan, '工号')
        dict_lvtong_total = specify_df_baoyuan(df_lvtong, '工号')
        """ 获取各人《抱怨》中装维无理失约 """
        dict_baoyuan_xiuzhang = specify_df_frequency(df_baoyuan, '工号', {'服务3': '超时未修复故障'})
        dict_baoyuan_zhuangji = specify_df_frequency(df_baoyuan, '工号', {'服务3': '未按预约时间上门'})
        """ 获取各人《绿通》中装维无理失约 """
        dict_lvtong_xiuzhang = specify_df_frequency(df_lvtong, '工号', {'服务类型': '修障问题-故障超时未修复（4小时）'})
        dict_lvtong_zhuangji = specify_df_frequency(df_lvtong, '工号', {'服务类型': '装移机问题-未按预约时间上门（4小时）'})
        """ 《抱怨》服务2：装/移机人员服务问题，维修人员服务问题 """
        dict_fuwutaidu_xiuzhang = specify_df_frequency(df_baoyuan, '工号', {'服务2': '装/移机人员服务问题'})
        dict_fuwutaidu_zhuangji = specify_df_frequency(df_baoyuan, '工号', {'服务2': '维修人员服务问题'})
        """ 获取光宽缓装虚假退单dict """
        print("\n************  处理（8/15）:光宽缓装虚假退单  ************")
        # file_tuidan = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\\10月光宽退单规范.xlsx'
        file_tuidan = f'{month}月光宽退单规范.xlsx'
        file_tuidan = file_tuidan if boolean==True else (input(f'\n请输入《{month}月光宽退单规范》文件名：（默认：{file_tuidan}）\n').strip() or file_tuidan)
        sheet_tuidan = '匹配清单'
        df_tuidan = get_specified_df(file_tuidan, sheet_tuidan, ['工号', '退单是否规范'])
        dict_tuidan = specify_df_frequency(df_tuidan, '工号', {'退单是否规范': '不规范'})
        """ 获取工信部dict """
        """ print("\n************  处理（9/15）：工信部单  ************")
        # file_gongxin = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\\工信部清单20191031.xlsx'
        file_gongxin = f'工信部清单{month}月.xlsx'
        file_gongxin = file_gongxin if boolean==True else (input(f'\n请输入《工信部清单{month}月》文件名：（默认：{file_gongxin}）\n').strip() or file_gongxin)
        sheet_gongxin = f'匹配清单'
        df_gongxin = get_specified_df(file_gongxin, sheet_gongxin, ['处理工号', '工单编号'])
        dict_gongxin = specify_df_frequency(df_gongxin, '处理工号') """
        """ 获取ivr故障虚假回单dict """
        print("\n************  处理（10/15）：ivr故障虚假回单  ************")
        # file_xujia_ivr_guzhang = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\\（IVR）故障虚假回单清单10月.xlsx'
        file_xujia_ivr_guzhang = f'（IVR）故障虚假回单清单{month}月.xlsx'
        file_xujia_ivr_guzhang = file_xujia_ivr_guzhang if boolean==True else (input(f'\n请输入《（IVR）故障虚假回单清单{month}月》文件名：（默认：{file_xujia_ivr_guzhang}）\n').strip() or file_xujia_ivr_guzhang)
        sheet_xujia_ivr_guzhang = '匹配清单'
        df_xujia_ivr_guzhang = get_specified_df(file_xujia_ivr_guzhang, sheet_xujia_ivr_guzhang, ['查修员工号', 'B、请问您的故障修复了吗？'])
        dict_xujia_ivr_guzhang = specify_df_frequency(df_xujia_ivr_guzhang, '查修员工号', {'B、请问您的故障修复了吗？': '未修复，请按3'})
        """ 获取ivr装机虚假回单dict """
        print("\n************  处理（11/15）：ivr装机虚假回单  ************")
        # file_xujia_ivr_zhuangji = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\\（IVR）装机虚假回单清单10月.xlsx'
        file_xujia_ivr_zhuangji = f'（IVR）装机虚假回单清单{month}月.xlsx'
        file_xujia_ivr_zhuangji = file_xujia_ivr_zhuangji if boolean==True else (input(f'\n请输入《（IVR）装机虚假回单清单{month}月》文件名：（默认：{file_xujia_ivr_zhuangji}）\n').strip() or file_xujia_ivr_zhuangji)
        sheet_xujia_ivr_zhuangji = '匹配清单'
        df_xujia_ivr_zhuangji = get_specified_df(file_xujia_ivr_zhuangji, sheet_xujia_ivr_zhuangji, ['装维人员工号', 'B、请问您的电信业务能正常使用吗？'])
        dict_xujia_ivr_zhuangji = specify_df_frequency_not_in_or(df_xujia_ivr_zhuangji, '装维人员工号',[{'B、请问您的电信业务能正常使用吗？': '能正常使用，请按2'}])
        """ 获取人工故障虚假回单dict """
        print("\n************  处理（12/15）：人工故障虚假回单  ************")
        # file_xujia_rengong_guzhang = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\\（人工）故障虚假回单清单10月.xlsx'
        file_xujia_rengong_guzhang = f'（人工）故障虚假回单清单{month}月.xlsx'
        file_xujia_rengong_guzhang =file_xujia_rengong_guzhang if boolean==True else (input(f'\n请输入《（人工）故障虚假回单清单{month}月》文件名：（默认：{file_xujia_rengong_guzhang}）\n').strip() or file_xujia_rengong_guzhang)
        sheet_xujia_rengong_guzhang = '匹配清单'
        df_xujia_rengong_guzhang = get_specified_df(file_xujia_rengong_guzhang, sheet_xujia_rengong_guzhang, ['查修员工号', 'D、请问维修人员有联系您处理过吗？'])
        dict_xujia_rengong_guzhang = specify_df_frequency(df_xujia_rengong_guzhang, '查修员工号', {'D、请问维修人员有联系您处理过吗？': '一直无人联系修障'})
        """ 获取人工装机虚假回单dict """
        print("\n************  处理（13/15）：人工装机虚假回单  ************")
        # file_xujia_rengong_zhuangji = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\\（人工）装机虚假回单清单10月.xlsx'
        file_xujia_rengong_zhuangji = f'（人工）装机虚假回单清单{month}月.xlsx'
        file_xujia_rengong_zhuangji = file_xujia_rengong_zhuangji if boolean==True else (input(f'\n请输入《（人工）故障虚假回单清单{month}月》文件名：（默认：{file_xujia_rengong_zhuangji}）\n').strip() or file_xujia_rengong_zhuangji)
        sheet_xujia_rengong_zhuangji = '匹配清单'
        column_xujia_rengong_zhuangji = ['装维人员工号', 'B、请问您的电信业务能正常使用吗？', 'E、请问是哪种情况不能使用？', 'F、请问是什么原因没当场安装好呢？']
        # df_xujia_rengong_zhuangji = get_specified_df(file_xujia_rengong_zhuangji, sheet_xujia_rengong_zhuangji, ['装维人员工号', 'B、请问您的电信业务能正常使用吗？'])
        df_xujia_rengong_zhuangji = get_specified_df(file_xujia_rengong_zhuangji, sheet_xujia_rengong_zhuangji, column_xujia_rengong_zhuangji)
        df_xujia_rengong_zhuangji = specify_df_in(df_xujia_rengong_zhuangji, {'E、请问是哪种情况不能使用？':'没当场装好，至今不能正常使用'})
        df_xujia_rengong_zhuangji = specify_df_not_include(df_xujia_rengong_zhuangji, {'B、请问您的电信业务能正常使用吗？': '能正常使用'})
        dict_xujia_rengong_zhuangji = specify_df_frequency_not_in_or(df_xujia_rengong_zhuangji, '装维人员工号',[{'F、请问是什么原因没当场安装好呢？': '我不需要安装电视机顶盒'}])
        """ 修障满意度 """
        print("\n************  处理（14/15）：修障满意度  ************")
        # file_manyi_xiuzhang = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\\修障服务测评清单（含IVR&人工）10月.xlsx'
        file_manyi_xiuzhang = f'修障服务测评清单（含IVR&人工）{month}月.xlsx'
        file_manyi_xiuzhang = file_manyi_xiuzhang if boolean==True else (input(f'\n请输入《修障服务测评清单（含IVR&人工）{month}月》文件名：（默认：{file_manyi_xiuzhang}）\n').strip() or file_manyi_xiuzhang)
        sheet_manyi_xiuzhang = '匹配清单'
        df_manyi_xiuzhang = get_specified_df(file_manyi_xiuzhang, sheet_manyi_xiuzhang, ['查修员工号', '满意度', '是否剔除', '测评结果来源', '产品类型'])
        df_manyi_xiuzhang = specify_df_in_or(df_manyi_xiuzhang, [{'测评结果来源':'10000号IVR'}, {'测评结果来源':'10000号人工'}, {'测评结果来源':'广东公司微信'}])
        df_manyi_xiuzhang = specify_df_in_or(df_manyi_xiuzhang, [{'产品类型':'光宽'}, {'产品类型':'铜宽'}])
        df_manyi_xiuzhang = df_manyi_xiuzhang.loc[~df_manyi_xiuzhang['是否剔除'].isin(['剔除'])]
        dict_manyi_xiuzhang_feichang = specify_df_frequency(df_manyi_xiuzhang, '查修员工号', {'满意度':'非常满意'})
        dict_manyi_xiuzhang_manyi = specify_df_frequency(df_manyi_xiuzhang, '查修员工号', {'满意度':'满意'})
        dict_manyi_xiuzhang_bumanyi = specify_df_frequency(df_manyi_xiuzhang, '查修员工号', {'满意度':'不满意'})  
        # print(dict_manyi_xiuzhang)
        """ 装机满意度 """
        print("\n************  处理（15/15）：装机满意度  ************")
        # file_manyi_zhuangji = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\\装机服务测评清单（含IVR&人工）10月.xlsx'
        file_manyi_zhuangji = f'装机服务测评清单（含IVR&人工）{month}月.xlsx'
        file_manyi_zhuangji = file_manyi_zhuangji if boolean==True else (input(f'\n请输入《装机服务测评清单（含IVR&人工）{month}月》文件名：（默认：{file_manyi_zhuangji}）\n').strip() or file_manyi_zhuangji)
        sheet_manyi_zhuangji = '匹配清单'
        df_manyi_zhuangji = get_specified_df(file_manyi_zhuangji, sheet_manyi_zhuangji, ['装维人员工号', '满意度', '是否剔除', '测评结果来源', '产品类型'])
        df_manyi_zhuangji = specify_df_in_or(df_manyi_zhuangji, [{'测评结果来源':'10000号IVR'}, {'测评结果来源':'10000号人工'}, {'测评结果来源':'广东公司微信'}])
        df_manyi_zhuangji = specify_df_in_or(df_manyi_zhuangji, [{'产品类型':'光宽'}, {'产品类型':'铜宽'}])
        df_manyi_zhuangji = df_manyi_zhuangji.loc[~df_manyi_zhuangji['是否剔除'].isin(['剔除'])]
        dict_manyi_zhuangji_feichang = specify_df_frequency(df_manyi_zhuangji, '装维人员工号', {'满意度':'非常满意'})
        dict_manyi_zhuangji_manyi = specify_df_frequency(df_manyi_zhuangji, '装维人员工号', {'满意度':'满意'})
        dict_manyi_zhuangji_bumanyi = specify_df_frequency(df_manyi_zhuangji, '装维人员工号', {'满意度':'不满意'})
        for i in dict_name:
            dict_name[i]['装移机'] = dict_zhuang.get(i, {0:0})[0]
            dict_name[i]['修障'] = dict_xiu.get(i, {0:0})[0]
            dict_name[i]['光衰整治'] = 0
            dict_name[i]['合计'] = dict_name[i]['装移机'] + dict_name[i]['修障'] + dict_name[i]['光衰整治']
            dict_name[i]['合计（日均8，20工作日）'] = 0 if dict_name[i]['合计']/20 <= 8 else 1
            dict_name[i]['抱怨量'] = dict_baoyuan_total.get(i, {0:0})[0] + dict_lvtong_total.get(i, {0:0})[0]
            dict_name[i]['非常满意'] = dict_manyi_xiuzhang_feichang.get(i, {0:0})[0] + dict_manyi_zhuangji_feichang.get(i, {0:0})[0]
            dict_name[i]['零抱怨奖金'] = 500 if dict_name[i]['抱怨量']==0 and dict_name[i]['合计（日均8，20工作日）'] > 0 else 0
            dict_name[i]['满意'] = dict_manyi_xiuzhang_manyi.get(i, {0:0})[0] + dict_manyi_zhuangji_manyi.get(i, {0:0})[0]
            dict_name[i]['不满意'] = dict_manyi_xiuzhang_bumanyi.get(i, {0:0})[0] + dict_manyi_zhuangji_bumanyi.get(i, {0:0})[0]
            dict_name[i]['满意奖金'] = dict_name[i]['满意'] * 0 + dict_name[i]['非常满意'] * 5
            """ 好像装、维无理失约放一起了，装机无理失约均为0? """
            dict_name[i]['装机无理失约'] = 0
            dict_name[i]['装维无理失约'] = dict_baoyuan_xiuzhang.get(i, {0:0})[0] + dict_baoyuan_zhuangji.get(i, {0:0})[0] + dict_lvtong_xiuzhang.get(i, {0:0})[0] + dict_lvtong_zhuangji.get(i, {0:0})[0]
            dict_name[i]['装机零失约奖金（暂停）'] = 0
            dict_name[i]['表扬（暂停）'] = 0
            dict_name[i]['表扬奖金'] = 0
            dict_name[i]['奖金'] = dict_name[i]['装机零失约奖金（暂停）'] + dict_name[i]['零抱怨奖金'] + dict_name[i]['满意奖金'] + dict_name[i]['表扬奖金']
            dict_name[i]['服务态度'] = dict_fuwutaidu_xiuzhang.get(i, {0:0})[0] + dict_fuwutaidu_zhuangji.get(i, {0:0})[0]
            dict_name[i]['虚假回单'] = dict_xujia_ivr_guzhang.get(i, {0:0})[0] + dict_xujia_ivr_zhuangji.get(i, {0:0})[0] + dict_xujia_rengong_guzhang.get(i, {0:0})[0] + dict_xujia_rengong_zhuangji.get(i, {0:0})[0]
            # dict_name[i]['工信'] = dict_gongxin.get(i, {0:0})[0]
            dict_name[i]['工信'] = 0
            dict_name[i]['光宽缓装虚假退单'] = dict_tuidan.get(i, {0:0})[0]
            dict_name[i]['催装催修≥4'] = dict_cuixiu.get(i, {0:0})[0] + dict_cuizhuang.get(i, {0:0})[0]
            dict_name[i]['扣罚分数'] = min(3*(dict_name[i]['装维无理失约'] + dict_name[i]['服务态度'] + dict_name[i]['虚假回单'] + dict_name[i]['光宽缓装虚假退单'] + dict_name[i]['催装催修≥4']) + 20*dict_name[i]['工信'] + 0.26*dict_name[i]['不满意'], 20)
            dict_name[i]['扣罚金额'] = min(dict_name[i]['扣罚分数']*38, 1000)
            dict_name[i]['扣罚次数'] = dict_name[i]['装维无理失约'] + dict_name[i]['服务态度'] + dict_name[i]['虚假回单'] + dict_name[i]['光宽缓装虚假退单'] + dict_name[i]['催装催修≥4'] + dict_name[i]['工信'] + dict_name[i]['不满意']
            dict_name[i]['奖励（单月扣罚3次不奖励）'] = 0 if dict_name[i]['扣罚次数'] > 3 else 1        
        df = pd.DataFrame.from_dict(dict_name, orient='index')
        # print(df)
        temp_excel_file = f'{os.path.dirname(os.path.realpath(sys.argv[0]))}\装维人员服务奖惩通报（四奖四罚）中间表：人员产能表（{month}月）{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
        df.to_excel(excel_writer = temp_excel_file, index = True)
        print("\n************  已完成  ************")
        print(f'\n保存地址{temp_excel_file}\n总耗时{time.time() - t_start}秒')
    except Exception as e:
        print(e)
    input(f'\n按回车离开')

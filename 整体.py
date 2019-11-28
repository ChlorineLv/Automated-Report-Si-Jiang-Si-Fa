# -*- coding: UTF-8 -*-
'''
:Version: Python 3.7.4
:Date: 2019-11-14 09:53:16
:LastEditors: ChlorineLv@outlook.com
:LastEditTime: 2019-11-28 16:28:32
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
    print(f'正在处理“人员名单”: {excel_file}')
    """ 只保留下面的列 """
    df = pd.DataFrame(pd.read_excel(excel_file))[['综调登录工号', '姓名', '手机', '分公司', '单位名称', '人员类别', '岗位类型', '在职状态']]
    """ 字段内容筛选 """
    df = df.loc[df['岗位类型'] == '外线施工岗'].loc[df['在职状态'] == '在职']
    """ 将手机号完整记录，不再是科学记数法 """
    df['手机'] = df['手机'].astype('int64')
    """ 工号通过lambda函数去除首部的0 """
    df['综调登录工号'] = df['综调登录工号'].map(lambda x: str(x)[1:] if str(x).startswith('0') else str(x))
    """ 工号列变index """
    df.set_index('综调登录工号', inplace=True)
    # print(df)
    """ 以index作为dict的key """
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


def get_zhuang_dict():
    '''
    :description: 以工号为key的各人装机次数
    :param {type} :
    :return: {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    csv_file = f'{os.path.dirname(__file__)}\表1装移机工单一览表.csv'
    # csv_file = input(f'请输入《装移机工单一览表》文件名：（默认：{csv_file}）\n').strip() or csv_file
    print(f'正在处理“装移机工单量”: {csv_file}')
    df = pd.DataFrame(pd.read_csv(csv_file, encoding='gbk'))[['施工类型','处理人工号']]
    df['处理人工号'] = df['处理人工号'].map(lambda x: str(x)[1:] if str(x).startswith('0') else str(x))
    """ 取反，取施工类型不包括拆机、其他、移拆机 """
    df = df[~df['施工类型'].isin(['拆机'])]
    df = df[~df['施工类型'].isin(['其他'])]
    df = df[~df['施工类型'].isin(['移拆'])]
    # print(df)
    """ 以处理人工号计算频次 """
    df = df.groupby(by='处理人工号').size().to_frame()
    """ 先在前面转为dataframe再转为dict """
    return df.to_dict(orient='index')


def get_xiu_dict():
    '''
    ### description: 以工号为key的各人修障次数
    ### param {type} 
    ### return: dict 
    ### example: {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    csv_file = f'{os.path.dirname(__file__)}\表2障碍用户申告一览表.csv'
    # csv_file = input(f'请输入《表2障碍用户申告一览表》文件名：（默认：{csv_file}）\n').strip() or csv_file
    print(f'正在处理“修障工单量”: {csv_file}')
    df = pd.DataFrame(pd.read_csv(csv_file, encoding='gbk'))[['部门分局','修理员工号']]
    df['修理员工号'] = df['修理员工号'].map(lambda x: str(x)[1:] if str(x).startswith('0') else str(x))
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
    print(f'正在处理:{excel_file}')
    if column_list != 0:
        df = pd.DataFrame(pd.read_excel(excel_file, sheet_name=excel_sheet, dtype={column_list[0]: str}))[column_list]
    else:
        df = pd.DataFrame(pd.read_excel(excel_file, sheet_name=excel_sheet, dtype={column_list[0]: str}))
    df[column_list[0]] = df[column_list[0]].map(lambda x: str(x)[1:] if str(x).startswith('0') else str(x))
    # print(df)
    return df


def specify_df_cui(df, c_name, c_judge):
    '''
    :description: 筛选‘催x次数’/‘前台催x次数’≥4次，且录音为空的
    :param df_get {dataframe} : dataframe
    :param c_name {string} : 工号
    :param c_judge {string} : "催装"、"催修"
    :return: dict
    '''
    df = df.fillna(0)
    """ 字段内容筛选‘催x次数’/‘前台催x次数’≥4次，且录音为空的 """
    df = df.loc[df[c_judge] >= 4]
    df = df.loc[df['录音开始时间'] == 0]
    """ 以‘处理人工号’/‘修理员工号’计算频次 """
    df = df.groupby(by=c_name).size().to_frame()
    print(c_name, df)
    """ 先在前面转为dataframe再转为dict """
    return df.to_dict(orient='index')


def specify_df_baoyuan(df_get, c_name):
    '''
    :description: 统计抱怨/绿通dataframe中各人抱怨量
    :param df_get {dataframe} :
    :param c_name {str} : 工号字段具体名称
    :return: dict {'工号1': {0: 次数1}, '工号2': {0: 次数2}, '工号3': {0: 次数3}, ......}
    '''
    # 以处理人工号计算频次
    df = df_get.groupby(by=c_name).size().to_frame()
    return df.to_dict(orient='index')

    
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


if __name__ == "__main__":
    t_start = time.time()
    """ 获取名单 """
    dict_name = get_name_list()
    """ 获取装移工作量 """
    dict_zhuang = get_zhuang_dict()
    """ 获取修障工作量 """
    dict_xiu = get_xiu_dict()
    """ 获取催装大于4次的单数 """
    file_cuizhuang = f'{os.path.dirname(__file__)}\催装催修清单10月.xlsx'
    # file_cuizhuang = input(f'请输入《催装催修清单10月》文件名：(默认：{file_cuizhuang})\n').strip() or file_cuizhuang
    sheet_cuizhuang = '催装'
    # sheet_cuizhuang = input(f'请输入sheet名：（默认{sheet_cuizhuang})\n').strip() or sheet_cuizhuang
    df_cuizhuang = get_specified_df(file_cuizhuang, sheet_cuizhuang, ['处理人工号', '前台催装次数', '录音开始时间'])
    dict_cuizhuang = specify_df_cui(df_cuizhuang, '处理人工号', '前台催装次数')
    """ 获取催修大于4次的单数 """
    file_cuixiu = f'{os.path.dirname(__file__)}\催装催修清单10月.xlsx'
    # file_cuixiu = input(f'请输入《催装催修清单10月》文件名：(默认：{file_cuixiu})\n').strip() or file_cuixiu
    sheet_cuixiu = '催修'
    # sheet_cuixiu = input(f'请输入sheet名：（默认{sheet_cuixiu})\n').strip() or sheet_cuixiu
    df_cuixiu = get_specified_df(file_cuixiu, sheet_cuixiu, ['修理员工号', '前台催修次数', '录音开始时间'])
    dict_cuixiu = specify_df_cui(df_cuixiu, '修理员工号', '前台催修次数')
    """ 获取《抱怨》的dataframe """
    file_baoyuan = f'{os.path.dirname(__file__)}\抱怨清单统计2019年10月.xlsx'
    # file_baoyuan = input(f'请输入《抱怨清单统计2019年10月》文件名：（默认：{file_baoyuan}）\n').strip() or file_baoyuan
    df_baoyuan = get_specified_df(file_baoyuan, '清单', ['工号', '服务2', '服务3'])
    """ 获取《绿通》的dataframe """
    file_lvtong = f'{os.path.dirname(__file__)}\绿通单清单2019年10月.xlsx'
    # file_lvtong = input(f'请输入《抱怨清单统计2019年10月》文件名：（默认：{file_lvtong}）\n').strip() or file_lvtong
    df_lvtong = get_specified_df(file_lvtong, '1', ['处理人工号','服务类型'])
    """ 获取各人《抱怨》总量和《绿通》总量 """
    dict_baoyuan_total = specify_df_baoyuan(df_baoyuan, '工号')
    dict_lvtong_total = specify_df_baoyuan(df_lvtong, '处理人工号')
    """ 获取各人《抱怨》中装维无理失约 """
    dict_baoyuan_xiuzhang = specify_df_frequency(df_baoyuan, '工号', {'服务3': '超时未修复故障'})
    dict_baoyuan_zhuangji = specify_df_frequency(df_baoyuan, '工号', {'服务3': '未按预约时间上门'})
    """ 获取各人《绿通》中装维无理失约 """
    dict_lvtong_xiuzhang = specify_df_frequency(df_lvtong, '处理人工号', {'服务类型': '修障问题-故障超时未修复（4小时）'})
    dict_lvtong_zhuangji = specify_df_frequency(df_lvtong, '处理人工号', {'服务类型': '装移机问题-未按预约时间上门（4小时）'})
    """ 《抱怨》服务2：装/移机人员服务问题，维修人员服务问题 """
    dict_fuwutaidu_xiuzhang = specify_df_frequency(df_baoyuan, '工号', {'服务2': '装/移机人员服务问题'})
    dict_fuwutaidu_zhuangji = specify_df_frequency(df_baoyuan, '工号', {'服务2': '维修人员服务问题'})
    """ 获取光宽缓装虚假退单dict """
    file_tuidan = f'{os.path.dirname(__file__)}\\10月光宽退单规范.xlsx'
    # file_tuidan = input(f'请输入《10月光宽退单规范》文件名：（默认：{file_tuidan}）\n').strip() or file_tuidan
    df_tuidan = get_specified_df(file_tuidan, 'Sheet2', ['工号', '退单是否规范'])
    dict_tuidan = specify_df_frequency(df_tuidan, '工号', {'退单是否规范': '不规范'})
    """ 获取工信部dict """
    file_gongxin = f'{os.path.dirname(__file__)}\\工信部清单20191031.xlsx'
    # file_gongxin = input(f'请输入《工信部清单20191031》文件名：（默认：{file_gongxin}）\n').strip() or file_gongxin
    df_gongxin = get_specified_df(file_gongxin, '10月', ['处理工号', '工单编号'])
    dict_gongxin = specify_df_frequency(df_gongxin, '处理工号')
    """ 获取ivr故障虚假回单dict """
    file_xujia_ivr_guzhang = f'{os.path.dirname(__file__)}\\（IVR）故障虚假回单清单2019年10月.xlsx'
    # file_xujia_ivr_guzhang = input(f'请输入《（IVR）故障虚假回单清单2019年10月》文件名：（默认：{file_xujia_ivr_guzhang}）\n').strip() or file_xujia_ivr_guzhang
    sheet_xujia_ivr_guzhang = '2019-10-08工单明细'
    # sheet_xujia_ivr_guzhang = input(f'请输入sheet名：（默认{sheet_xujia_ivr_guzhang})\n').strip() or sheet_xujia_ivr_guzhang
    df_xujia_ivr_guzhang = get_specified_df(file_xujia_ivr_guzhang, sheet_xujia_ivr_guzhang, ['查修员工号', 'B、请问您的故障修复了吗？'])
    dict_xujia_ivr_guzhang = specify_df_frequency(df_xujia_ivr_guzhang, '查修员工号', {'B、请问您的故障修复了吗？': '未修复，请按3'})
    """ 获取ivr装机虚假回单dict """
    file_xujia_ivr_zhuangji = f'{os.path.dirname(__file__)}\\（IVR）装机虚假回单清单2019年10月.xlsx'
    # file_xujia_ivr_zhuangji = input(f'请输入《（IVR）装机虚假回单清单2019年10月》文件名：（默认：{file_xujia_ivr_zhuangji}）\n').strip() or file_xujia_ivr_zhuangji
    sheet_xujia_ivr_zhuangji = '2019-10-08工单明细'
    # sheet_xujia_ivr_zhuangji = input(f'请输入sheet名：（默认{sheet_xujia_ivr_zhuangji})\n').strip() or sheet_xujia_ivr_zhuangji
    df_xujia_ivr_zhuangji = get_specified_df(file_xujia_ivr_zhuangji, sheet_xujia_ivr_zhuangji, ['装维人员工号', 'B、请问您的电信业务能正常使用吗？'])
    dict_xujia_ivr_zhuangji = specify_df_frequency_not_in_or(df_xujia_ivr_zhuangji, '装维人员工号',[{'B、请问您的电信业务能正常使用吗？': '能正常使用，请按2'}])
    """ 获取人工故障虚假回单dict """
    file_xujia_rengong_guzhang = f'{os.path.dirname(__file__)}\\（人工）故障虚假回单清单2019年10月.xlsx'
    # file_xujia_rengong_guzhang = input(f'请输入《（人工）故障虚假回单清单2019年10月》文件名：（默认：{file_xujia_rengong_guzhang}）\n').strip() or file_xujia_rengong_guzhang
    sheet_xujia_rengong_guzhang = 'Sheet1'
    # sheet_xujia_rengong_guzhang = input(f'请输入sheet名：（默认{sheet_xujia_rengong_guzhang})\n').strip() or sheet_xujia_rengong_guzhang
    df_xujia_rengong_guzhang = get_specified_df(file_xujia_rengong_guzhang, sheet_xujia_rengong_guzhang, ['查修员工号', 'D、请问维修人员有联系您处理过吗？'])
    dict_xujia_rengong_guzhang = specify_df_frequency(df_xujia_rengong_guzhang, '查修员工号', {'D、请问维修人员有联系您处理过吗？': '一直无人联系修障'})
    """ 获取人工装机虚假回单dict """
    file_xujia_rengong_zhuangji = f'{os.path.dirname(__file__)}\\（人工）装机虚假回单清单2019年10月.xlsx'
    # file_xujia_rengong_zhuangji = input(f'请输入《（人工）故障虚假回单清单2019年10月》文件名：（默认：{file_xujia_rengong_zhuangji}）\n').strip() or file_xujia_rengong_zhuangji
    sheet_xujia_rengong_zhuangji = 'Sheet1'
    # sheet_xujia_rengong_zhuangji = input(f'请输入sheet名：（默认{sheet_xujia_rengong_zhuangji})\n').strip() or sheet_xujia_rengong_zhuangji
    column_xujia_rengong_zhuangji = ['装维人员工号', 'B、请问您的电信业务能正常使用吗？', 'E、请问是哪种情况不能使用？', 'F、请问是什么原因没当场安装好呢？']
    # df_xujia_rengong_zhuangji = get_specified_df(file_xujia_rengong_zhuangji, sheet_xujia_rengong_zhuangji, ['装维人员工号', 'B、请问您的电信业务能正常使用吗？'])
    df_xujia_rengong_zhuangji = get_specified_df(file_xujia_rengong_zhuangji, sheet_xujia_rengong_zhuangji, column_xujia_rengong_zhuangji)
    dict_xujia_rengong_zhuangji = specify_df_frequency_not_in_or(df_xujia_rengong_zhuangji, '装维人员工号',[{'B、请问您的电信业务能正常使用吗？': '能正常使用'}])

    n = 0
    for i in dict_name:
        dict_name[i]['装移机'] = dict_zhuang.get(i, {0:0})[0]
        dict_name[i]['修障'] = dict_xiu.get(i, {0:0})[0]
        dict_name[i]['光衰整治'] = 0
        dict_name[i]['合计'] = dict_name[i]['装移机'] + dict_name[i]['修障'] + dict_name[i]['光衰整治']
        dict_name[i]['合计（日均8，20工作日）'] = 0 if dict_name[i]['合计']/20 < 8 else 1
        dict_name[i]['非常满意'] = 0
        dict_name[i]['满意'] = 0
        dict_name[i]['不满意'] = 0
        dict_name[i]['满意奖金'] = dict_name[i]['满意'] * 0 + dict_name[i]['非常满意'] * 5
        dict_name[i]['表扬（暂停）'] = 0
        dict_name[i]['表扬奖金'] = 0
        dict_name[i]['奖金'] = 0
        dict_name[i]['抱怨量'] = dict_baoyuan_total.get(i, {0:0})[0] + dict_lvtong_total.get(i, {0:0})[0]
        """ 好像装、维无理失约放一起了，装机无理失约均为0? """
        dict_name[i]['装机无理失约'] = 0
        dict_name[i]['装维无理失约'] = dict_baoyuan_xiuzhang.get(i, {0:0})[0] + dict_baoyuan_zhuangji.get(i, {0:0})[0] + dict_lvtong_xiuzhang.get(i, {0:0})[0] + dict_lvtong_zhuangji.get(i, {0:0})[0]
        dict_name[i]['装机零失约奖金（暂停）'] = 0
        dict_name[i]['零抱怨奖金'] = 1 if dict_name[i]['装机无理失约'] == 0 and dict_name[i]['装维无理失约'] == 0 and dict_name[i]['合计（日均8，20工作日）'] >= 0 else 0
        dict_name[i]['服务态度'] = dict_fuwutaidu_xiuzhang.get(i, {0:0})[0] + dict_fuwutaidu_zhuangji.get(i, {0:0})[0]
        dict_name[i]['虚假回单'] = dict_xujia_ivr_guzhang.get(i, {0:0})[0] + dict_xujia_ivr_zhuangji.get(i, {0:0})[0] + dict_xujia_rengong_guzhang.get(i, {0:0})[0] + dict_xujia_rengong_zhuangji.get(i, {0:0})[0]
        dict_name[i]['工信'] = dict_gongxin.get(i, {0:0})[0]
        dict_name[i]['光宽缓装虚假退单'] = dict_tuidan.get(i, {0:0})[0]
        dict_name[i]['催装催修≥4'] = dict_cuixiu.get(i, {0:0})[0] + dict_cuizhuang.get(i, {0:0})[0]
        dict_name[i]['扣罚分数'] = min(3*(dict_name[i]['装维无理失约'] + dict_name[i]['服务态度'] + dict_name[i]['虚假回单'] + dict_name[i]['光宽缓装虚假退单'] + dict_name[i]['催装催修≥4']) + 20*dict_name[i]['工信'] + 0.26*dict_name[i]['不满意'], 20)
        dict_name[i]['扣罚金额'] = min(dict_name[i]['扣罚分数']*38, 1000)
        dict_name[i]['扣罚次数'] = dict_name[i]['装维无理失约'] + dict_name[i]['服务态度'] + dict_name[i]['虚假回单'] + dict_name[i]['光宽缓装虚假退单'] + dict_name[i]['催装催修≥4'] + dict_name[i]['工信'] + dict_name[i]['不满意']
        dict_name[i]['奖励（单月扣罚3次不奖励）'] = 0 if dict_name[i]['扣罚次数'] > 3 else 1

        n+=1
                
    df = pd.DataFrame.from_dict(dict_name, orient='index')
    # print(df)
    temp_excel_file = f'{os.path.dirname(__file__)}\中间表：人员产能表：{time.strftime("%Y-%m-%d", time.localtime())}.xlsx'
    df.to_excel(excel_writer = temp_excel_file, index = True)
    print(f'已完成，保存地址{temp_excel_file}\n总耗时{time.time() - t_start}秒')
# -*- coding: utf-8 -*-
"""
Created on Fri Oct 23 17:54:57 2020

@author: klmtl

此程序解决从云南电网公司安全监察系统中导出的安全监督计划中提取安监人员
每人到位次数。
"""

import pandas as pd
from pandas.api.types import CategoricalDtype


class getExcelFile:
    '''
    依据给定的文件全路径及excel文件中的sheet名称，读取excel文件到DataFrame;
    通过getDf()方法获取该DataFrame
    '''

    def __init__(self, path, sheet):
        ex_file = pd.ExcelFile(path, engine='openpyxl')
        self.df = ex_file.parse(sheet)

    def getDf(self):
        return self.df


class saveExcelFile:
    '''
    将给定的DataFrame存储到给定的全路径文件名xlsx文件中。
    '''

    def __init__(self, df, path):
        df.to_excel(path, index=False, engine='openpyxl')


class countTimes:
    '''
    依据给定的计划DataFrame和人员DataFrame，生产到位次数DataFrame
    '''

    def __init__(self, df_plan, df_name):
        # 将null值填充为空字符串,不然后续筛选将出错。
        df_plan[['检查负责人', '检查人']] = df_plan[['检查负责人', '检查人']].fillna('')
        # 将人员表姓名列中的null值删除
        df_name = df_name.dropna(subset=['姓名'])
        # 将人员表姓名列中的待定值删除
        df_name = df_name[df_name['姓名'] != '待定']
        # 将人员表姓名转换成list为后续判断做准备
        df_name_list = df_name['姓名'].tolist()
        # 将人员表保留需要信息
        df_name = df_name[['姓名', '单位', '部门', '岗位']]
        # 将人员表合并成以|为分隔的字符串为str.contains多条件判断做准备
        contains_text = '|'.join(df_name_list)
        # 从计划表中筛选出'检查负责人'中为安监人员及'检查人'中有安监人员的所有记录
        df_plan_sf = df_plan[(df_plan['检查负责人'].isin(df_name_list)) | (df_plan['检查人'].str.contains(contains_text))]
        path1 = r'D:\JGY\600-Data\002-in输入文件\02-work工作\01-system工作系统数据\监督计划 (过程).xlsx'
        df_plan_sf.to_excel(path1, index=False, engine='openpyxl')
        # 设置记录字典
        dt_count = {}
        # 对安监人员进行循环
        for name in df_name_list:
            # 每个循环清空df
            df_n_name = pd.DataFrame()
            # 筛选出含有name的记录筛及选出含有name所在单位的记录
            df_n_name = df_plan_sf[((df_plan_sf['检查负责人'] == name) | (df_plan_sf['检查人'].str.contains(name))) & (
                df_plan_sf[r'检查部门/单位'].str.contains(''.join(df_name[df_name['姓名'] == name]['单位'].values)))]
            # 输出每个人记录
            path2 = r'D:\JGY\600-Data\002-in输入文件\02-work工作\01-system工作系统数据\监督计划 (' + ''.join(
                df_name[df_name['姓名'] == name]['单位'].values) + name + ').xlsx'
            df_n_name.to_excel(path2, index=False, engine='openpyxl')

            # 记录个数
            dt_count[name] = len(df_n_name)

        # 将获得的字典统计值转换成df类型
        df_count = pd.DataFrame.from_dict(dt_count, orient='index')
        # 将名字列从index回归到column
        df_count = df_count.reset_index()
        # 将列重新命名
        df_count.columns = ['姓名', '到位次数']
        # 和人员表进行merge
        df_count = pd.merge(df_name, df_count, on='姓名').sort_values(by='到位次数', ascending=False)
        # 对index进行重置
        df_count = df_count.reset_index(drop=True)
        self.df_count = df_count

    def getDfCount(self):
        return self.df_count

    def pandas_sort(self, df, columns_name, list_sort):
        for i in range(len(columns_name)):
            cat_order = CategoricalDtype(
                list_sort[i],
                ordered=True
            )

            df[columns_name[i]] = df[columns_name[i]].astype(cat_order)

        df = df.sort_values(columns_name, axis=0, ascending=[True, True])

        df.reset_index(drop=True, inplace=True)
        return df


if __name__ == '__main__':
    def main():
        # 更改此监督计划文件路径就可以生产新的到位次数
        ex_supervision_plan = r'D:\JGY\600-Data\002-in输入文件\02-work工作\01-system工作系统数据\001-安全监督计划\监督计划 (系统导出2.1-2.28).xlsx'
        # 调取getExcelFile类获取excel文件转成df
        df_plan = getExcelFile(ex_supervision_plan, 'Sheet1')

        # 更改此安监人员情况
        ex_supervision_people = r'D:\JGY\600-Data\002-in输入文件\02-work工作\03-document工作文档\昆明供电局领导及安监人员.xlsx'
        # 调取getExcelFile类获取excel文件转成df
        df_name = getExcelFile(ex_supervision_people, 'Sheet1')
        # 调取getExcelFile类获取excel文件转成df
        df_list = getExcelFile(ex_supervision_people, 'Sheet2')
        df_list_sort = df_list.getDf()
        df_list_sort_d = df_list_sort[~df_list_sort['单位'].isnull()]['单位']

        # 调取countTimes类获安监人员到位次数
        ct = countTimes(df_plan.getDf(), df_name.getDf())
        ct_sort = ct.pandas_sort(ct.getDfCount(), ['单位', '姓名'], [df_list_sort_d, df_list.getDf()['姓名']])
        # 更改存储路径
        ex_save = r'D:\JGY\600-Data\003-out输出文件\02-work工作\03-document工作文档\领导及监督人员到位.xlsx'
        saveExcelFile(ct_sort, ex_save)


    main()

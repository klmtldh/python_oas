import pandas as pd
from pandas.api.types import CategoricalDtype
import numpy as np
import matplotlib.pyplot as plt
from pylab import mpl
import os

import pymysql
from sqlalchemy import create_engine

import datetime
import calendar
from dateutil import rrule
from dateutil.relativedelta import relativedelta

import kl_class


# 将安全监督管理平台人员信息导出表转化为工作数据库可以导入的excel文件
def ydsfs_toexcel(read_path, save_path, read_sheet_name='Sheet1', save_sheet_name='Sheet1'):
    xl = pd.ExcelFile(read_path, engine='openpyxl')
    df = xl.parse(read_sheet_name)
    df.fillna('', inplace=True)  # 将nan转为空字符，不然后续出错
    df1 = df[df['所队(县区级单位)'] != '']
    df1_1 = df1.copy()  # 进行复制避免警告
    df1_1.loc[:, '单位字符'] = '/' + df1_1.loc[:, '分子公司'] + \
                           '/' + df1_1.loc[:, '地市级单位'] + '/' + \
                           df1_1.loc[:, '所队(县区级单位)'] + '/'
    df2 = df[df['所队(县区级单位)'] == '']
    df2_1 = df2.copy()
    df2_1.loc[:, '单位字符'] = '/' + df2_1.loc[:, '分子公司'] + \
                           '/' + df2_1.loc[:, '地市级单位'] + '/'
    df = pd.concat([df1_1, df2_1])
    df3 = df.copy()
    for i in range(len(df)):
        df3.loc[i, '部门'] = df3.loc[i, '部门'].replace(str(df3.loc[i, '单位字符']), '')
    df4 = df3['部门'].str.split('/', expand=True)
    del df3['序号']
    del df3['部门']
    del df3['单位字符']
    df4.columns = ['部门', '科室']
    df = df3.join(df4, how='left')
    df = df[[
        '分子公司',
        '地市级单位',
        '所队(县区级单位)',
        '4A账号',
        '姓名',
        '部门',
        '科室',
        '岗位',
        '性别',
        '出生年月(2020-1-1)',
        '毕业院校(初始)',
        '所学专业(初始)',
        '初始学历',
        '后续学历',
        '职称',
        '安监技术专家等级',
        '安监技能专家等级',
        '安监技术能力等级',
        '持证情况1-安风体系审核员级别',
        '持证情况2-注册安全工程师',
        '持证情况3-其他',
        '安监专业之外从事过的主要专业',
        '参加工作起始年(2020-1-1)',
        '从事安监专业起始时间(2020-1-1)',
        '获得奖励,专利等情况',
        '备注'
    ]]
    df.to_excel(save_path, sheet_name=save_sheet_name, index=False, engine='openpyxl')
    return df


# 将工作数据库导出的数据转换成安全监督管理平台人员信息可以导入的格式
def excel_toydsfs(read_path, save_path, read_sheet_name='Sheet1', save_sheet_name='Sheet1'):
    xl = pd.ExcelFile(read_path, engine='openpyxl')
    df = xl.parse(read_sheet_name)
    df.fillna('', inplace=True)
    df1 = df[df['所队(县区级单位)'] != '']
    df1_1 = df1.copy()
    df1_1.loc[:, '单位字符'] = '/' + df1_1.loc[:, '分子公司'] + '/' + \
                           df1_1.loc[:, '地市级单位'] + '/' + \
                           df1_1.loc[:, '所队(县区级单位)'] + '/' + df1_1.loc[:, '部门']
    df2 = df[df['所队(县区级单位)'] == '']
    df2_1 = df2.copy()
    df2_1.loc[:, '单位字符'] = '/' + df2_1.loc[:, '分子公司'] + '/' + \
                           df2_1.loc[:, '地市级单位'] + '/' + df2_1.loc[:, '部门']
    df = pd.concat([df1_1, df2_1])

    df_1 = df[df['科室'] != '']
    df_1_1 = df_1.copy()
    df_1_1.loc[:, '部门'] = df_1_1.loc[:, '单位字符'] + '/' + df_1_1.loc[:, '科室']
    df_2 = df[df['科室'] == '']
    df_2_1 = df_2.copy()
    df_2_1.loc[:, '部门'] = df_2_1.loc[:, '单位字符']
    df = pd.concat([df_1_1, df_2_1])

    del df['科室']
    del df['单位字符']
    df.reset_index(drop=True, inplace=True)
    df.loc[:, '序号'] = df.loc[:, '部门'].index + 1
    df = df[['序号',
             '分子公司',
             '地市级单位',
             '所队(县区级单位)',
             '4A账号',
             '姓名',
             '部门',
             '岗位',
             ' 性别',
             '出生年月(2020-1-1)',
             '毕业院校(初始)',
             '所学专业(初始)',
             '初始学历',
             '后续学历',
             '职称',
             '安监技术专家等级',
             '安监技能专家等级',
             '安监技术能力等级',
             '持证情况1-安风体系审核员级别',
             '持证情况2-注册安全工程师',
             '持证情况3-其他',
             '安监专业之外从事过的主要专业',
             '参加工作起始年(2020-1-1)',
             '从事安监专业起始时间(2020-1-1)',
             '获得奖励,专利等情况',
             '备注'
             ]
            ]
    df.to_excel(save_path, sheet_name=save_sheet_name, index=False, engine='openpyxl')
    return df


# 更新数据库
def update_excel(excel_name, sheet_name, table_name):
    xlsql = kl_class.xlsxMysql(excel_name, sheet_name)
    xlsql.to_mysql(table_name)


def analysis_null(sql):
    dfsql = kl_class.dfMysql()
    df = dfsql.select_mysql(sql)
    columns_list = list(df)
    pdf = kl_class.parsingDf()
    df_null, df_null_text = pdf.null_item(df)
    pdf.df_fig(df_null, 'bar')
    return len(df), columns_list, df_null, df_null_text


def analysis_item(sql, item):
    dfsql = kl_class.dfMysql()
    df = dfsql.select_mysql(sql)
    df = df.applymap(str)
    pdf = kl_class.parsingDf()
    df_item, df_item_text = pdf.df_item(df, item)
    pdf.df_fig(df_item, 'bar')
    return df_item, df_item_text


# 计算年龄，空值用当天日期代替。
def pandas_age(column):
    now_year = datetime.date.today().year
    column.fillna(datetime.date.today(), inplace=True)
    column = pd.to_datetime(column)
    age = now_year - column.dt.year
    return age


def redundant_vacancies(sql_person, sql_value):
    # 连接数据库，读取人员信息和人员配置信息
    dfsql = kl_class.dfMysql()
    df_person = dfsql.select_mysql(sql_person)
    df_person['年龄'] = pandas_age(df_person['出生年月(2020-1-1)'])
    df_person['工龄'] = pandas_age(df_person['参加工作起始年(2020-1-1)'])

    df_value = dfsql.select_mysql(sql_value)

    # 统一人员配置数据
    df_value = df_value[['单位', '机构名称', '科室', '岗位', '定编']]
    df_value.rename(columns={'机构名称': '部门'}, inplace=True)

    # 解决告警问题，一是要copy，二是要用loc引用
    df_value1 = df_value.copy()
    df_value1.loc[:, ['定编']] = df_value1.loc[:, ['定编']].astype('int')

    # 对人员表进行透视
    for i in range(len(df_person)):
        if pd.isna(df_person.loc[i, '所队(县区级单位)']):
            df_person.loc[i, '单位'] = df_person.loc[i, '地市级单位']
        else:
            df_person.loc[i, '单位'] = df_person.loc[i, '所队(县区级单位)']
    df_person.fillna('无', inplace=True)
    df_person_talbe = pd.pivot_table(df_person,
                                     index=['单位', '部门', '科室', '岗位'],
                                     values=['姓名'],
                                     aggfunc=len,
                                     fill_value=0,
                                     margins=True,
                                     # margins_name='总数'
                                     )
    df_person_talbe.reset_index(inplace=True)
    df_person_talbe.columns = ['单位', '部门', '科室', '岗位', '在岗数']
    df_value1.fillna('无', inplace=True)
    df_person_talbe.fillna('无', inplace=True)
    df_table = pd.merge(df_value1, df_person_talbe, how='left', on=['单位', '部门', '科室', '岗位'])
    df_table.fillna(0, inplace=True)
    df_table['缺员'] = df_table['定编'] - df_table['在岗数']

    text = "昆明供电局安全监管部（应急指挥中心）定编{0}人，\
实际{1}人，缺员{2}人。".format(
        df_table['定编'].sum(),
        df_table['在岗数'].sum(),
        df_table['缺员'].sum()
    )
    pdf = kl_class.parsingDf()

    # 党员
    # df_person_party=df_person['政治面貌'].value_counts()
    # df_person_party = kl_class.sort_list(df_person_party, ['党员', '群众'])
    # text_person_party = pdf.series_tex(df_person_party, '名')

    # 性别
    df_person_sex = df_person['性别'].value_counts()
    df_person_sex = kl_class.sort_list(df_person_sex, ['男', '女'])
    text_person_sex = pdf.series_tex(df_person_sex, '名')

    # 年龄
    df_person_age = pd.cut(df_person['年龄'].astype('int'), bins=[0, 30, 40, 50, 100],
                           labels=["30岁以下", "30-39岁", "40-49岁", "50岁以上"])
    df_person_age_count = df_person_age.value_counts()
    df_person_age_count = kl_class.sort_list(df_person_age_count, ["30岁以下", "30-39岁", "40-49岁", "50岁以上"])

    text_person_age_count = pdf.series_tex(df_person_age_count, '名')

    #工龄
    df_person_wage = pd.cut(df_person['工龄'].astype('int'), bins=[0, 5, 10, 20, 30, 50],
                           labels=["5年以下", "5-9年", "10-19年","20-29年", "30年以上"])
    df_person_wage_count = df_person_wage.value_counts()
    df_person_wage_count = kl_class.sort_list(df_person_wage_count, ["5年以下", "5-9年", "10-19年","20-29年", "30年以上"])

    text_person_wage_count = pdf.series_tex(df_person_wage_count, '名')

    # 文化程度
    text_list = []
    list_columns = ['初始学历', '所学专业(初始)', '职称']
    s_list = ['男', '女',
              '党员', '群众',
              '研究生', '本科', '大专',
              '高级工程师', '工程师', '助理工程师',
              '经理', '副经理', '主管',
              '安全监察专责(A)', '安全监察专责(B)', '风险体系管理专责(A)', '风险体系管理专责(B)', '应急管理专责',
              '班长', '安全监察员']
    for item in list_columns:
        df_person_x = df_person[item].value_counts()
        df_person_x = kl_class.sort_list(df_person_x, s_list)
        text_person_x = pdf.series_tex(df_person_x, '名')
        text_list.append(text_person_x)

    text = text + text_person_sex + text_person_age_count + '工龄：' + text_person_wage_count + ''.join(text_list)

    df_person = df_person[['姓名', '单位', '部门', '科室', '岗位', '性别', '年龄']]
    df_person.fillna('无', inplace=True)
    df_person = kl_class.sort_list(df_person, s_list, '岗位')
    list_d = [
        '昆明供电局',
        '官渡供电局',
        '五华供电局',
        '盘龙供电局',
        '西山供电局',
        '呈贡供电局',
        '东川供电局',
        '昆明安宁供电局',
        '云南电网有限责任公司昆明晋宁供电局',
        '云南电网有限责任公司昆明宜良供电局',
        '云南电网有限责任公司昆明富民供电局',
        '云南电网有限责任公司昆明嵩明供电局',
        '云南电网有限责任公司昆明石林供电局',
        '云南电网有限责任公司昆明寻甸供电局',
        '云南电网有限责任公司昆明禄劝供电局',
    ]

    list_b = ['安全监管部（应急指挥中心）',
              '系统运行部',
              '科技创新及数字化中心',
              '规划建设管理中心（质监分站）',
              '供电服务中心',
              '综合服务中心（涉电分中心）',
              '物流服务中心',
              '带电作业中心',
              '通信管理所',
              '输电管理所',
              '变电运行一所',
              '变电运行二所',
              '变电修试所',
              '安全监管部',
              ]

    list_k = [
        '无',
        '安全监察科',
        '应急与保供电管理科',
        '安全督查大队',
        '安全监察班',
        '综合室',
        '安生室',
        '计划安全室',
        '质量安全室',
        '科技室',
    ]
    list_sort = [list_d, list_b, list_k]
    column = ['单位', '部门', '科室']
    for i in range(len(column)):
        cat_order = CategoricalDtype(
            list_sort[i],
            ordered=True
        )

        df_person[column[i]] = df_person[column[i]].astype(cat_order)

    df_person = df_person.sort_values(column, axis=0, ascending=[True, True, True])

    df_person.reset_index(drop=True, inplace=True)

    return df_table, text, df_person


def write_docx(sql, sql_value):
    doc = kl_class.oas_docx()
    doc.hd('昆明供电局安全监管人员情况', font_size=22)
    correct_tex = "昆明供电局安全监管人员情况共有{0}条数据，以下数据不完善：{1}" \
        .format(analysis_null(sql)[0], analysis_null(sql)[3])
    doc.par('一、数据总体情况', bold=True)
    doc.par(correct_tex)

    doc.par('二、人员总体情况', bold=True)

    df, text_person, df_person = redundant_vacancies(sql, sql_value)
    doc.par(text_person)
    doc.par('附件：安监部人员在岗情况表')
    doc.tb(len(df) + 1, len(list(df)))
    col_list = list(df)
    for i in range(len(df) + 1):
        for j in range(len(list(df))):
            if i == 0:
                doc.tb_cell(col_list[j], i, j, f_name='黑体')
            else:
                doc.tb_cell(str(df.iloc[i - 1, j]), i, j)

    doc.par('附件：安监部人员信息表')

    doc.tb(len(df_person) + 1, len(list(df_person)))
    col_list = list(df_person)
    for i in range(len(df_person) + 1):

        for j in range(len(list(df_person))):

            if i == 0:
                doc.tb_cell(col_list[j], i, j, f_name='黑体')

            else:
                doc.tb_cell(str(df_person.iloc[i - 1, j]), i, j)

    if kl_class.get_pc_name() == 'LAPTOP-HF9P6H1P':
        # 个人
        doc.save_docx(r'D:\JGY\600-Data\003-out输出文件\02-work工作\03-document工作文档\昆明供电局安监部人员信息（全局）.docx')
    else:
        # 单位
        doc.save_docx(r'Z:\100内部管理\101组织管理\1012部门人员信息\昆明供电局安监部人员信息（全局）.docx')


if __name__ == '__main__':
    #转换系统数据
    read_path = r'D:\JGY\600-Data\002-in输入文件\02-work工作\01-system工作系统数据\安全监管人员信息.xlsx'
    save_path = r'D:\JGY\600-Data\003-out输出文件\02-work工作\02-database工作数据库\安全监管人员信息.xlsx'
    ydsfs_toexcel(read_path, save_path)
    #excel_toydsfs(read_path, save_path)
    # 更新数据库
    excel_name = save_path
    sheet_name = 'Sheet1'
    table_name = '昆明供电局安监人员信息'
    update_excel(excel_name, sheet_name, table_name)
    # sheet_name = '安监部岗位设置'
    # table_name = '安监部岗位设置'
    # update_excel(excel_name, sheet_name, table_name)
    # 统一查询库
    sql = "select * from 昆明供电局安监人员信息"
    sql_value = "select * from 昆明供电局安监人员岗位设置"
    # x = redundant_vacancies(sql, sql_value)
    # print(x)
    # 写入文档
    write_docx(sql, sql_value)

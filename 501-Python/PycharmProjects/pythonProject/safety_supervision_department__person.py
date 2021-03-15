import pandas as pd
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

def update_excel(excel_name,sheet_name,table_name):
    xlsql=kl_class.xlsxMysql(excel_name,sheet_name)
    xlsql.to_mysql(table_name)
def analysis_null(sql):
    dfsql = kl_class.dfMysql()
    df = dfsql.select_mysql(sql)
    columns_list = list(df)
    pdf = kl_class.parsingDf()
    df_null, df_null_text = pdf.null_item(df)
    pdf.df_fig(df_null,'bar')
    return len(df),columns_list,df_null,df_null_text
def analysis_item(sql,item):
    dfsql = kl_class.dfMysql()
    df = dfsql.select_mysql(sql)
    df=df.applymap(str)
    pdf = kl_class.parsingDf()
    df_item, df_item_text = pdf.df_item(df,item)
    pdf.df_fig(df_item,'bar')
    return df_item,df_item_text


def redundant_vacancies(sql_person,sql_value):
    #连接数据库，读取人员信息和人员配置信息
    dfsql = kl_class.dfMysql()
    df_person = dfsql.select_mysql(sql_person)
    df_value = dfsql.select_mysql(sql_value)

    #统一人员配置数据
    df_value = df_value[['科室或班组名称', '岗位名称', '职数']]
    df_value.rename(columns={'科室或班组名称': '科室', '岗位名称': '工作岗位'}, inplace=True)
    #df_value['职数']=df_value['职数'].astype('int')
    df_value1 = df_value.copy()
    df_value1.loc[:, ['职数']] = df_value1.loc[:, ['职数']].astype('int')

    #对人员表进行透视
    df_person_talbe = pd.pivot_table(df_person,
                                     index=['科室', '工作岗位'],
                                     columns=['岗位状态'],
                                     values=['姓名'],
                                     aggfunc=len,
                                     fill_value=0,
                                     margins=True,
                                     margins_name='总数'
                                     )
    df_person_talbe.reset_index(inplace=True)
    df_person_talbe.columns=['科室', '工作岗位', '借入', '借出', '在岗', '总数']

    df_table = pd.merge(df_value1, df_person_talbe, how='left', on=['科室', '工作岗位'])
    df_table['缺员'] = df_table['职数']-df_table['总数']



    text = "昆明供电局安全监管部（应急指挥中心）定员{0}人，\
实际{1}人，缺员{2}人,其中：在岗{3}人，借入{4}人，借出{5}人。".format(
        df_table['职数'].sum(),
        df_table['总数'].sum(),
        df_table['缺员'].sum(),
        df_table['在岗'].sum(),
        df_table['借入'].sum(),
        df_table['借出'].sum()
    )
    pdf=kl_class.parsingDf()

    #党员
    df_person_party=df_person['政治面貌'].value_counts()
    df_person_party = kl_class.sort_list(df_person_party, ['党员', '群众'])
    text_person_party = pdf.series_tex(df_person_party, '名')

    #性别
    df_person_sex=df_person['性别'].value_counts()
    df_person_sex = kl_class.sort_list(df_person_sex, ['男', '女'])
    text_person_sex = pdf.series_tex(df_person_sex, '名')

    #年龄
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

    #文化程度
    text_list=[]
    list_columns=['民族', '文化程度', '职称']
    s_list=['男', '女',
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

    text=text+text_person_party\
         +text_person_sex+text_person_age_count\
         +'工龄：'+text_person_wage_count+''.join(text_list)

    df_person = df_person[['姓名', '科室', '工作岗位', '兼任职务', '联系电话', '政治面貌', '工号', '性别', '岗位状态']]
    df_person.fillna('无',inplace=True)
    df_person = kl_class.sort_list(df_person, s_list,'工作岗位')
    df_person.reset_index(drop=True,inplace=True)


    return df_table,text,df_person


def write_docx(sql,sql_value):
    doc = kl_class.oas_docx()
    doc.hd('昆明供电局安全监管部（应急指挥中心）人员情况', font_size=22)
    correct_tex = "昆明供电局安全监管部（应急指挥中心）人员情况共有{0}条数据，以下数据不完善：{1}"\
        .format(analysis_null(sql)[0], analysis_null(sql)[3])
    doc.par('一、数据总体情况', bold=True)
    doc.par(correct_tex)

    doc.par('二、人员总体情况', bold=True)

    df,text_person,df_person=redundant_vacancies(sql, sql_value)
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



    if kl_class.get_pc_name()=='LAPTOP-HF9P6H1P':
        # 个人
        doc.save_docx(r'D:\JGY\600-Data\003-out输出文件\工作\昆明供电局安监部人员信息.docx')
    else:
        #单位
        doc.save_docx(r'Z:\100内部管理\101组织管理\1012部门人员信息\昆明供电局安监部人员信息.docx')

if __name__ == '__main__':
    #更新数据库
    # excel_name = r'Z:\100内部管理\101组织管理\1012部门人员信息\昆明供电局安监部人员信息.xlsx'
    # sheet_name = '安监部人员信息'
    # table_name = '安监部人员信息'
    # update_excel(excel_name,sheet_name,table_name)
    # sheet_name = '安监部岗位设置'
    # table_name = '安监部岗位设置'
    # update_excel(excel_name, sheet_name, table_name)
    #统一查询库
    sql = "select * from 安监部人员信息 WHERE 岗位状态!='调离'"
    sql_value = "select * from 安监部岗位设置"
    # x = redundant_vacancies(sql, sql_value)
    # print(x)
    #写入文档
    write_docx(sql, sql_value)



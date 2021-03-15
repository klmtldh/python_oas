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
def analysis_noarrange(sql):
    dfsql = kl_class.dfMysql()
    df = dfsql.select_mysql(sql)
    df_section_noarrange = df[(df['开始时间'].isnull())|(df['结束时间'].isnull())]
    df_section_noarrange = df_section_noarrange.fillna('未填写')
    pdf = kl_class.parsingDf()
    return pdf.df_iter( df_section_noarrange )



def analysis_section(sql,flag,s_date,e_date=datetime.date.today()):
    dfsql = kl_class.dfMysql()
    df = dfsql.select_mysql(sql)
    start_date = pd.to_datetime(s_date).date()
    end_date = pd.to_datetime(e_date).date()
    df_section_end = df[(df['开始时间'] >= start_date)
                        & (df['结束时间'] <= end_date)
                        & (df['是否完成整改'] == flag)
                    ]
    df_section_end = df_section_end.fillna('未填写')
    pdf = kl_class.parsingDf()
    return pdf.df_iter( df_section_end )

def analysis_name(sql,name,flag,s_date,e_date=datetime.date.today()):
    dfsql = kl_class.dfMysql()
    df = dfsql.select_mysql(sql)
    start_date = pd.to_datetime(s_date).date()
    end_date = pd.to_datetime(e_date).date()
    df_section_end = df[(df['开始时间'] >= start_date)
                        & (df['结束时间'] <= end_date)
                        & (df['是否完成整改'] == flag)
                        & (df['整改责任人'] == name)
                       ]
    df_section_end = df_section_end.fillna('未填写')
    pdf = kl_class.parsingDf()
    return pdf.df_iter( df_section_end )



def write_docx(sql):
    doc = kl_class.oas_docx()
    doc.hd('立行立改问题库', font_size=22)
    correct_tex = "昆明供电局安全监管部（应急指挥中心）立行立改问题库共有{0}条数据，以下数据不完善：{1}"\
        .format(analysis_null(sql)[0], analysis_null(sql)[3])
    doc.par('一、总体情况')
    doc.par(correct_tex)
    if kl_class.get_pc_name()=='LAPTOP-HF9P6H1P':
        #个人
        doc.pic(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png', wth=6)
        os.unlink(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png')
    else:
        #单位
        doc.pic(r'D:\out\fig.png', wth=6)
        os.unlink(r'D:\out\fig.png')
    columns_list=analysis_null(sql)[1]
    columns_list.remove('问题编号')
    columns_list.remove('发现问题')
    columns_list.remove('整改措施')
    columns_list.remove('发现时间')
    columns_list.remove('开始时间')
    columns_list.remove('结束时间')
    columns_list.remove('整改情况及进度')
    columns_list.remove('备注')

    for item in columns_list:
        try:
            df_item,df_item_text=analysis_item(sql, item)
            doc.par('{0}：{1}'.format(item,df_item_text))
            if kl_class.get_pc_name() == 'LAPTOP-HF9P6H1P':
                # 个人
                doc.pic(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png', wth=6)
                os.unlink(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png')
            else:
                # 单位
                doc.pic(r'D:\out\fig.png', wth=6)
                os.unlink(r'D:\out\fig.png')
        except Exception as e:
            print(e)
    doc.par('二、未安排事项')
    noarrange_list=analysis_noarrange(sql)
    i=1
    for item in noarrange_list:
        doc.par('（{0}）未安排{0}'.format(kl_class.digital_to_chinese(i)))
        x=item.split('\n')
        for y in x:
            doc.par(y)
        i+=1


    doc.par('三、周期事项事项')
    analysis_list=analysis_section(sql,'否','2020-1-1')
    #print(analysis_list)
    i=1
    for item in analysis_list:
        doc.par('（{0}）本周事项{0}'.format(kl_class.digital_to_chinese(i)))
        x=item.split('\n')
        for y in x:
            doc.par(y)
        i+=1

    doc.par('四、个人未完成事项')
    dfsql = kl_class.dfMysql()
    df = dfsql.select_mysql(sql)
    duty_name = list(set(df['整改责任人'].tolist()))
    duty_name.sort()

    j=1
    for name in duty_name:
        analysis_name_list = analysis_name(sql, name, '否', '2020-1-1')
        doc.par('（{0}）{1}'.format(kl_class.digital_to_chinese(j), name))
        if analysis_name_list==[]:
            doc.par('无异常')
        else:
            i = 1
            for item in analysis_name_list:
                doc.par('{0}.未完成事项{0}'.format(i))
                x = item.split('\n')
                for y in x:
                    doc.par(y)
                i += 1

        j+=1
    if kl_class.get_pc_name()=='LAPTOP-HF9P6H1P':
        # 个人
        doc.save_docx(r'D:\JGY\600-Data\003-out输出文件\工作\立行立改问题库.docx')
    else:
        #单位
        doc.save_docx(r'Z:\200业务管理\206问题收集与处置\2061纠正与预防问题库\立行立改问题记录表\立行立改问题库.docx')

if __name__ == '__main__':
    # excel_name = r'Z:\200业务管理\206问题收集与处置\2061纠正与预防问题库\立行立改问题记录表\立行立改问题库.xlsx'
    # sheet_name = '立行立改问题库'
    # table_name = '立行立改问题库'
    # update_excel(excel_name,sheet_name,table_name)
    # 统一查询库
    sql = "select * from 立行立改问题库"
    write_docx(sql)



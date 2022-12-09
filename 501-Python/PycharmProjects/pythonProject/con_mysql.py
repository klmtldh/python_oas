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

import oas_docx
import pandas_mysql
import xlsx_mysql

_MAPPING = (u'零', u'一', u'二', u'三', u'四', u'五', u'六', u'七', u'八', u'九', u'十', u'十一', u'十二', u'十三', u'十四', u'十五', u'十六', u'十七',u'十八', u'十九')
_P0 = (u'', u'十', u'百', u'千',)
_S4 = 10 ** 4
def _to_chinese4(num):
    assert (0 <= num and num < _S4)
    if num < 20:
        return _MAPPING[num]
    else:
        lst = []
        while num >= 10:
            lst.append(num % 10)
            num = num / 10
        lst.append(num)
        c = len(lst)  # 位数
        result = u''

        for idx, val in enumerate(lst):
            val = int(val)
            if val != 0:
                result += _P0[idx] + _MAPPING[val]
                if idx < c - 1 and lst[idx + 1] == 0:
                    result += u'零'
        return result[::-1]





def series_tex(sr, unit):
    text = ''
    for i in range(len(sr)):
        if i != len(sr) - 1:
            text += sr.index[i] + str(sr[i]) + unit + '，'
        else:
            text += sr.index[i] + str(sr[i]) + unit + '。'

    return text



if __name__ == "__main__":
    excel_name=r'Z:\200业务管理\206问题收集与处置\2061纠正与预防问题库\立行立改问题记录表\立行立改问题库.xlsx'
    sheet_name='立行立改问题库'
    table_name = '立行立改问题库'
    xsql=xlsx_mysql.xlsxMysql(excel_name,sheet_name)
    xsql.to_mysql(table_name)
    dfsql = pandas_mysql.dfMysql()
    mpl.rcParams['font.sans-serif'] = ['Microsoft YaHei']
    # 解决图表负号显示不正确问题
    plt.rcParams['axes.unicode_minus'] = False
    m_correct = oas_docx.oas_docx()

    sql = "select * from 立行立改问题库"
    df = dfsql.select_mysql( sql)
    correct_null = df.isnull().sum()

    correct_null.sort_values(ascending = False,inplace = True)
    print(correct_null)
    correct_null_text = series_tex(correct_null, '项')
    correct_doing = df['是否完成整改'].value_counts()
    correct_doing_text = series_tex(correct_doing, '项')
    correct_duty = df['整改责任人'].value_counts()
    correct_duty_text = series_tex(correct_duty, '项')


    # 文档标题
    m_correct.hd('立行立改问题库', font_size=22)
    # 立行立改问题库基本情况
    correct_tex = "昆明供电局安全监管部（应急指挥中心）立行立改问题库共有{0}条数据，\
以下数据不完善：{1}".format(len(df), correct_null_text)

    # 统计周期数据
    start_date = pd.to_datetime('2020-1-1')
    end_date = pd.to_datetime('2021-12-31')

    duty_name = list(set(df['整改责任人'].tolist()))
    duty_name.sort()

    today = datetime.date.today()
    df_section_noarrange = df[(df['开始时间'].isnull())|(df['结束时间'].isnull())]
    df_section_noarrange=df_section_noarrange.fillna('未填写')


    df_section_end = df[(df['开始时间'] >= start_date)
                    & (df['结束时间'] <= end_date)
                    & (df['是否完成整改']=='是')
                    & (df['整改责任人'].isin(duty_name))
                    ]
    df_section_end = df_section_end.fillna('未填写')
    # 总体

    m_correct.par('一、总体情况')
    m_correct.par(correct_tex)
    ax = correct_null.plot(kind='bar',
                           figsize=(8, 5))

    fig = ax.get_figure()
    fig.savefig(r'D:\out\fig.png')
    m_correct.pic(r'D:\out\fig.png', wth=6)
    os.unlink(r'D:\out\fig.png')
    plt.clf()
    m_correct.par('是否完成：{0}'.format(correct_doing_text))

    ax = correct_doing.plot(kind='pie',
                            autopct='%.1f%%',
                            figsize=(8, 5))

    fig = ax.get_figure()
    fig.savefig(r'D:\out\fig.png')
    m_correct.pic(r'D:\out\fig.png', wth=6)
    os.unlink(r'D:\out\fig.png')
    plt.clf()
    m_correct.par('整改责任人：{0}'.format(correct_duty_text))

    ax = correct_duty.plot(kind='bar',
                           figsize=(8, 5))

    fig = ax.get_figure()
    fig.savefig(r'D:\out\fig.png')
    m_correct.pic(r'D:\out\fig.png', wth=6)
    os.unlink(r'D:\out\fig.png')
    plt.clf()
    m_correct.par('二、未安排事项')
    count = 0
    for tup in zip(df_section_noarrange['发现问题'],
                   df_section_noarrange['整改措施'],
                   df_section_noarrange['开始时间'],
                   df_section_noarrange['结束时间'],
                   df_section_noarrange['整改情况及进度'],
                   df_section_noarrange['整改责任人'],
                   ):
        m_correct.par('发现问题'+_to_chinese4(count+1)+'：'+tup[0])
        m_correct.par('整改措施' +_to_chinese4(count+1)+'：'+ tup[1])
        m_correct.par('开始时间' + _to_chinese4(count+1) + '：' + str(tup[2]))
        m_correct.par('结束时间' + _to_chinese4(count+1) + '：' + str(tup[3]))
        m_correct.par('整改情况及进度'+_to_chinese4(count+1)+'：'+tup[4])
        m_correct.par('整改责任人' + _to_chinese4(count + 1) + '：' + tup[5])
        m_correct.par('')
        count += 1

#个人
    m_correct.par('三、个人未完成数据')
    count=0
    for name in duty_name:
        m_correct.par('(' + _to_chinese4(count + 1) + '）' + name)
        print(name)
        df_section_doing = df[(
                                      (df['开始时间'] >= start_date)
                                      & (df['结束时间'] <= today)
                                      & (df['是否完成整改'] == '否')
                                      & (df['整改责任人']==name)
                              )

                              ]
        df_section_doing = df_section_doing.fillna('未填写')
        #print(df_section_doing)
        if len(df_section_doing)==0:
            m_correct.par('无异常项')
        else:
            count1=0
            for tup in zip(df_section_doing['发现问题'],
                           df_section_doing['整改措施'],
                           df_section_doing['开始时间'],
                           df_section_doing['结束时间'],
                           df_section_doing['整改情况及进度'],
                           df_section_doing['整改责任人'],
                           ):
                m_correct.par('发现问题' + _to_chinese4(count1 + 1) + '：' + tup[0])
                m_correct.par('整改措施' + _to_chinese4(count1 + 1) + '：' + tup[1])
                m_correct.par('开始时间' + _to_chinese4(count1 + 1) + '：' + str(tup[2]))
                m_correct.par('结束时间' + _to_chinese4(count1 + 1) + '：' + str(tup[3]))
                m_correct.par('整改情况及进度' + _to_chinese4(count1 + 1) + '：' + tup[4])
                m_correct.par('整改责任人' + _to_chinese4(count1 + 1) + '：' + tup[5])
                m_correct.par('')
                count1 +=1
        count += 1
    m_correct.par('四、统计周期正常推进情况')
    count = 0
    for tup in zip(df_section_end['发现问题'],
                   df_section_end['整改措施'],
                   df_section_end['开始时间'],
                   df_section_end['结束时间'],
                   df_section_end['整改情况及进度'],
                   df_section_end['整改责任人']):
        m_correct.par('发现问题' + _to_chinese4(count + 1) + '：' + tup[0])
        m_correct.par('整改措施' + _to_chinese4(count + 1) + '：' + tup[1])
        m_correct.par('开始时间' + _to_chinese4(count + 1) + '：' + str(tup[2]))
        m_correct.par('结束时间' + _to_chinese4(count + 1) + '：' + str(tup[3]))
        m_correct.par('整改情况及进度' + _to_chinese4(count + 1) + '：' + tup[4])
        m_correct.par('整改责任人' + _to_chinese4(count + 1) + '：' + tup[5])
        m_correct.par('')
        count += 1
    m_correct.save_docx(r'Z:\200业务管理\206问题收集与处置\2061纠正与预防问题库\立行立改问题记录表\立行立改问题库.docx')

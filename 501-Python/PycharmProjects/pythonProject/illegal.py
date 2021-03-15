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


def select_mysql(con, sql):
    engine = create_engine(con)
    select = sql
    df = pd.read_sql(select, engine)
    return df


def series_tex(sr, unit):
    text = ''
    for i in range(len(sr)):
        if i != len(sr) - 1:
            text += sr.index[i] + str(sr[i]) + unit + '，'
        else:
            text += sr.index[i] + str(sr[i]) + unit + '。'

    return text

def xlsx_mysql(con,name,sheet_name,table_name):
    engine = create_engine(con)
    db = pymysql.connect(host='localhost',
                         port=3306,
                         user='root',
                         passwd='kongling8167',
                         db='kmsf',
                         charset='utf8'
                         )
    df=pd.read_excel(name,sheet_name,engine='openpyxl')
    USER_TABLE_NAME = table_name
    try:
        # 执行SQL语句
        df.to_sql('temp', engine, if_exists='replace', index=False)  # 把新数据写入 temp 临时表
        connection = db.cursor()
        # 替换数据的语句
        args1 = f" REPLACE INTO {USER_TABLE_NAME} SELECT * FROM temp "
        connection.execute(args1)
        args2 = " DROP Table If Exists temp"  # 把临时表删除
        connection.execute(args2)
        # 提交到数据库执行
        db.commit()
        connection .close()
        db.close()

    except:
        # 发生错误时回滚
        db.rollback()
        # 关闭数据库连接
        connection .close()
        db.close()

def all_list():
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

if __name__ == "__main__":
    excel_name=r'D:\JGY\600-Data\002-in输入文件\工作\数据库\违章台账 .xlsx'
    sheet_name='Sheet1'
    table_name='违章台账'
    con = 'mysql+pymysql://root:kongling8167@localhost:3306/kmsf?charset=utf8'
    xlsx_mysql(con,excel_name,sheet_name,table_name)
    mpl.rcParams['font.sans-serif'] = ['Microsoft YaHei']
    # 解决图表负号显示不正确问题
    plt.rcParams['axes.unicode_minus'] = False
    m_illegal = oas_docx.oas_docx()

    sql = "select * from 违章台账"
    df = select_mysql(con, sql)
    illegal_null = df.isnull().sum()

    illegal_null.sort_values(ascending = False,inplace = True)

    illegal_null_text = series_tex(illegal_null, '项')
    illegal_level = df['违章等级'].value_counts()
    illegal_level_text = series_tex(illegal_level, '项')
    illegal_major = df['违章主体'].value_counts()
    illegal_major_text = series_tex(illegal_major, '项')
    illegal_expose = df['是否自主暴露'].value_counts()
    illegal_expose_text = series_tex(illegal_expose, '项')
    illegal_enterprise = df['是否改革后企业'].value_counts()
    illegal_enterprise_text = series_tex(illegal_enterprise, '项')
    # 文档标题
    m_illegal.hd('昆明供电局违章基础分析', font_size=22)
    # 立行立改问题库基本情况
    illegal_tex = "昆明供电局违章共有{0}条数据，\
以下数据不完善：{1}".format(len(df), illegal_null_text)

    # 总体

    m_illegal.par('一、总体情况')
    m_illegal.par(illegal_tex)
    ax = illegal_null.plot(kind='bar',
                           figsize=(8, 5))

    fig = ax.get_figure()
    fig.savefig(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png')
    m_illegal.pic(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png', wth=6)
    os.unlink(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png')
    plt.clf()
    m_illegal.par('违章等级：{0}'.format(illegal_level_text))

    ax = illegal_level.plot(kind='pie',
                            autopct='%.1f%%',
                            figsize=(8, 5))

    fig = ax.get_figure()
    fig.savefig(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png')
    m_illegal.pic(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png', wth=6)
    os.unlink(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png')
    plt.clf()
    m_illegal.par('违章主体：{0}'.format(illegal_major_text))

    ax = illegal_major.plot(kind='bar',
                           figsize=(8, 5))

    fig = ax.get_figure()
    fig.savefig(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png')
    m_illegal.pic(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png', wth=6)
    os.unlink(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png')
    plt.clf()

    m_illegal.par('是否自主暴露：{0}'.format(illegal_expose_text))

    ax = illegal_expose.plot(kind='bar',
                           figsize=(8, 5))

    fig = ax.get_figure()
    fig.savefig(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png')
    m_illegal.pic(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png', wth=6)
    os.unlink(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png')
    plt.clf()

    m_illegal.par('是否改革后企业：{0}'.format(illegal_enterprise_text))

    ax = illegal_enterprise.plot(kind='bar',
                           figsize=(8, 5))

    fig = ax.get_figure()
    fig.savefig(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png')
    m_illegal.pic(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png', wth=6)
    os.unlink(r'D:\JGY\600-Data\003-out输出文件\工作\fig.png')
    plt.clf()
    illegal_name=list(set(df['违章人'].tolist()))
    m_illegal.par('违章人员清单：{0}'.format('、'.join(illegal_name)))
    m_illegal.save_docx(r'D:\JGY\600-Data\003-out输出文件\工作\违章台账.docx')

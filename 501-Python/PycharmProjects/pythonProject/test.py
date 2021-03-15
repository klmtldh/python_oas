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

from pyecharts import options as opts
from pyecharts.components import Table
def table_base(header,row) -> Table:
    table = Table(page_title='未完成项',js_host=r"D:\JGY\600-Data\004-auxiliary辅助文件\\")
    headers = header
    rows = row
    table.add(headers, rows).set_global_opts(
        title_opts=opts.ComponentTitleOpts(title="Table")
    )
    return table


def to_echarts(sql):
    from pyecharts import options as opts
    from pyecharts.charts import Gauge
    dfsql = kl_class.dfMysql()
    df = dfsql.select_mysql(sql)
    #删除重复值
    df.drop_duplicates(inplace=True)
    #删除缺失值
    #print(len(df))
    df.dropna(axis=0, how='all', inplace=True)
    #print(len(df))
    #填充缺失值
    #df.fillna('未填写',inplace=True)
    #删除空格
    df['是否完成整改']=df['是否完成整改'].map(str.strip)
    start_date='2020-1-1'
    end_date = pd.to_datetime(datetime.date.today()).date()
    df= df[(df['开始时间'] >= pd.to_datetime(start_date).date())
                        & (df['结束时间'] <= end_date)
                        ]

    df_doing=df[df['是否完成整改']=='否']
    df_split = df_doing.to_dict(orient='split')
    tb = table_base(df_split['columns'], df_split['data'])
    tb.render("doing.html")

    zg=df['是否完成整改'].value_counts()
    rat=round(100*(zg['是']/len(df)),2)
    ajk_name=['李黎','陈虹伍','杨和俊','陈斌','何艳琪','郭晶晶','杨建平']
    yjk_name=['洪永健','和定繁','王玺','周浩然','高璐']
    df_split1=kl_class.split_record(df, '整改责任人', '、')

    df_ajk=df_split1[df_split1['整改责任人'].isin(ajk_name)]
    zg_ajk=df_ajk['是否完成整改'].value_counts()
    rat_ajk=round(100*(zg_ajk['是']/len(df_ajk)),2)
    df_yjk=df_split1[df_split1['整改责任人'].isin(yjk_name)]
    zg_yjk=df_yjk['是否完成整改'].value_counts()
    rat_yjk=round(100*(zg_yjk['是']/len(df_yjk)),2)




    #交叉表
    ct=pd.crosstab(index=df_split1['整改责任人'], columns=df_split1['是否完成整改'], margins=True)
    #ct.columns=[]
    ct['rat']=round(100*(ct['是']/ct['All']),2)
    ct.sort_values('rat',ascending=False,inplace=True)

    t=[]

    for column in ct.iteritems():
        if column[0]=='rat':

            for i in range(len(ct)):
                if column[1].index[i]!='All':
                    t.append(column[1].index[i]+str(column[1].values[i])+'%\n')

    tn=''.join(t)
    # a = df[''].str.split('、', expand=True).stack().value_counts()
    # print(a)
    text = "安全监管部完成率{0}%\n\n安全监察科完成率{1}%：\n\n应急管理科完成率{2}%：\n\n{3}".format(rat, rat_ajk, rat_yjk, tn)

    c = (
        # Gauge(init_opts=opts.InitOpts(width="1200px", height="600px"))#设置画布尺寸
        Gauge(init_opts=opts.InitOpts(
            js_host=r"D:\JGY\600-Data\004-auxiliary辅助文件\\",
            width="1200px", height="600px",
            renderer="RenderType.SVG",
            page_title="立行立改问题库",
            theme="white")
        )
            .add(
            series_name="安全监管部",
            data_pair=[("安全监管部", rat)],
            is_selected=True,
            min_=0,
            max_=100,
            split_number=10,
            axisline_opts=opts.AxisLineOpts(
                linestyle_opts=opts.LineStyleOpts(
                    color=[(0.3, "#FF0033"), (0.7, "#FF6600"), (1, "#009900")], width=30
                )
            ),

            radius="80%",
            start_angle=150,
            end_angle=30,
            is_clock_wise=True,
            title_label_opts=opts.GaugeTitleOpts(
                font_size=20,
                color="blue",
                font_family="Microsoft YaHei",
                offset_center=[0, "-55%"],
            ),
            detail_label_opts=opts.GaugeDetailOpts(
                is_show=True,
                background_color='red',
                border_width=0,
                border_color='blue',
                offset_center=[0, "-40%"],
                color="white",
                font_style="oblique",
                font_weight="bold",
                font_family="Microsoft YaHei",
                font_size=26,
                border_radius=120,
                padding=[4, 4, 4, 4],
                formatter="{value}%"
            ),
            pointer=opts.GaugePointerOpts(
                # 是否显示指针。
                is_show=True,

                # 指针长度，可以是绝对数值，也可以是相对于半径的半分比。
                length="80%",

                # 指针宽度。
                width=12,
            ),
        )
            .add(
            series_name="安全监察科",
            data_pair=[("安全监察科", rat_ajk)],
            is_selected=True,
            min_=0,
            max_=100,
            split_number=10,
            axisline_opts=opts.AxisLineOpts(
                linestyle_opts=opts.LineStyleOpts(
                    color=[(0.3, "#FF0033"), (0.7, "#FF6600"), (1, "#009900")], width=30
                )
            ),

            radius="80%",
            start_angle=265,
            end_angle=155,
            is_clock_wise=True,
            title_label_opts=opts.GaugeTitleOpts(
                font_size=20,
                color="blue",
                font_family="Microsoft YaHei",
                offset_center=[-100, "25%"],
            ),
            detail_label_opts=opts.GaugeDetailOpts(
                is_show=True,
                background_color='red',
                border_width=0,
                border_color='blue',
                offset_center=[-100, "40%"],
                color="white",
                font_style="oblique",
                font_weight="bold",
                font_family="Microsoft YaHei",
                font_size=26,
                border_radius=120,
                padding=[4, 4, 4, 4],
                formatter="{value}%"
            ),
            pointer=opts.GaugePointerOpts(
                # 是否显示指针。
                is_show=True,

                # 指针长度，可以是绝对数值，也可以是相对于半径的半分比。
                length="50%",

                # 指针宽度。
                width=8,
            ),
        )
            .add(
            series_name="应急管理科",
            data_pair=[("应急管理科", rat_yjk)],
            is_selected=True,
            min_=0,
            max_=100,
            split_number=10,
            axisline_opts=opts.AxisLineOpts(
                linestyle_opts=opts.LineStyleOpts(
                    color=[(0.3, "#FF0033"), (0.7, "#FF6600"), (1, "#009900")], width=30
                )
            ),
            radius="80%",
            start_angle=385,
            end_angle=275,
            is_clock_wise=True,
            title_label_opts=opts.GaugeTitleOpts(
                font_size=20,
                color="blue",
                font_family="Microsoft YaHei",
                offset_center=[100, "25%"],
            ),
            detail_label_opts=opts.GaugeDetailOpts(
                is_show=True,
                background_color='red',
                border_width=0,
                border_color='blue',
                offset_center=[100, "40%"],
                color="white",
                font_style="oblique",
                font_weight="bold",
                font_family="Microsoft YaHei",
                font_size=26,
                border_radius=120,
                padding=[4, 4, 4, 4],
                formatter="{value}%"
            ),
            pointer=opts.GaugePointerOpts(
                # 是否显示指针。
                is_show=True,

                # 指针长度，可以是绝对数值，也可以是相对于半径的半分比。
                length="50%",

                # 指针宽度。
                width=8,
            ),
        )
            .set_global_opts(
            title_opts=opts.TitleOpts(
                title="立行立改问题库完成率",
                title_link="https:www.163.com",
                title_target="self",
                subtitle=text,
                subtitle_link="doing.html",
                subtitle_target="blank",
                pos_left="left",

            )
        )
            .render("gauge_lxlxwtk.html")
    )
if __name__ == '__main__':
    if kl_class.get_pc_name()=='LAPTOP-HF9P6H1P':
        # 个人
        excel_name = r'D:\JGY\600-Data\002-in输入文件\工作\数据库\立行立改问题库.xlsx'
    else:
        #单位
        excel_name = r'Z:\200业务管理\206问题收集与处置\2061纠正与预防问题库\立行立改问题记录表\立行立改问题库.xlsx'

    sheet_name = '立行立改问题库'
    table_name = '立行立改问题库'
    update_excel(excel_name,sheet_name,table_name)
    # 统一查询库
    sql = "select * from 立行立改问题库"
    write_docx(sql)
    to_echarts(sql)



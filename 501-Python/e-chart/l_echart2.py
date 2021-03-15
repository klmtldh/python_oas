# -*- coding: utf-8 -*-
"""
Created on Wed Dec 30 11:39:19 2020

@author: Kong Ling

E-Mail:klmtldh@163.com

QQ:1170233

Wechat:klmtldh
"""

from pyecharts import options as opts
from pyecharts.charts import Bar, Grid, Line, Liquid, Page, Pie
from pyecharts.commons.utils import JsCode
from pyecharts.components import Table
from pyecharts.faker import Faker
import pandas as pd

def bar_datazoom_slider(xaxis,yaxis_name,yaxis,title_name) -> Bar:
   
    if isinstance(yaxis_name,str):
            c=(
        Bar()
        .add_xaxis(xaxis)
        .add_yaxis(yaxis_name,yaxis)
        .set_global_opts(
            title_opts=opts.TitleOpts(title=title_name),
            datazoom_opts=[opts.DataZoomOpts()],
            brush_opts=opts.BrushOpts(),
            toolbox_opts=opts.ToolboxOpts(),
            legend_opts=opts.LegendOpts(is_show=False),
            yaxis_opts=opts.AxisOpts(name="我是 Y 轴"),
            xaxis_opts=opts.AxisOpts(name="我是 X 轴"),
        )
    )
    elif isinstance(yaxis,list):
        c=Bar()
        c.add_xaxis(xaxis)
        for i in range(len(yaxis_name)):
            c.add_yaxis(yaxis_name[i],yaxis[i], gap="0%",category_gap="10%",is_selected=True)
        c.set_global_opts(
            title_opts=opts.TitleOpts(title=title_name),
            datazoom_opts=[opts.DataZoomOpts(), opts.DataZoomOpts(type_="inside")],
            #datazoom_opts=opts.DataZoomOpts(orient="vertical"),
            brush_opts=opts.BrushOpts(),
            toolbox_opts=opts.ToolboxOpts(),
            legend_opts=opts.LegendOpts(is_show=False),
            yaxis_opts=opts.AxisOpts(name="我是 Y 轴"),
            xaxis_opts=opts.AxisOpts(name="我是 X 轴"),
        )
        c.set_series_opts(
        label_opts=opts.LabelOpts(is_show=True),
        markpoint_opts=opts.MarkPointOpts(
            data=[
                opts.MarkPointItem(type_="max", name="最大值"),
                opts.MarkPointItem(type_="min", name="最小值"),
                opts.MarkPointItem(type_="average", name="平均值"),
            ]
        )
        )              
  
    
    return c


def table_base(header,row) -> Table:
    table = Table()
    headers = header
    rows = row
    table.add(headers, rows).set_global_opts(
        title_opts=opts.ComponentTitleOpts(title="Table")
    )
    return table
name=r'D:\JGY\600-Data\002-in输入文件\个人\孔令手机通信录.xlsx'
#df=pd.read_excel(name,engine='openpyxl')
#df1=df[['姓名','手机','单位','职务']].head(5)
#df_split=df1.to_dict(orient='split')
#tb=table_base(df_split['columns'],df_split['data'])
#tb.render_notebook()

xaxis=['2019年','2020年','2021年']
yaxis_name=['测试1','测试2','测试3']
#yaxis_name='测试'
yaxis=[[1,2,3],[4,5,6],[7,8,9]]
b=bar_datazoom_slider(xaxis,yaxis_name,yaxis,'孔令编制')
b.render_notebook()
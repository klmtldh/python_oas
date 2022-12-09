# -*- coding: utf-8 -*-
"""
Created on Sat Dec 19 22:42:01 2020

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
import os as os
from bs4 import BeautifulSoup

def bar_datazoom_slider() -> Bar:
    c = (
        Bar()
        .add_xaxis(Faker.days_attrs)
        .add_yaxis("商家A", Faker.days_values)
        .set_global_opts(
            title_opts=opts.TitleOpts(title="Bar-DataZoom（slider-水平）"),
            datazoom_opts=[opts.DataZoomOpts()],
        )
    )
    return c


def line_markpoint() -> Line:
    c = (
        Line()
        .add_xaxis(Faker.choose())
        .add_yaxis(
            "商家A",
            Faker.values(),
            markpoint_opts=opts.MarkPointOpts(data=[opts.MarkPointItem(type_="min")]),
        )
        .add_yaxis(
            "商家B",
            Faker.values(),
            markpoint_opts=opts.MarkPointOpts(data=[opts.MarkPointItem(type_="max")]),
        )
        .set_global_opts(title_opts=opts.TitleOpts(title="Line-MarkPoint"))
    )
    return c


def pie_rosetype() -> Pie:
    v = Faker.choose()
    c = (
        Pie()
        .add(
            "",
            [list(z) for z in zip(v, Faker.values())],
            radius=["30%", "75%"],
            center=["25%", "50%"],
            rosetype="radius",
            label_opts=opts.LabelOpts(is_show=False),
        )
        .add(
            "",
            [list(z) for z in zip(v, Faker.values())],
            radius=["30%", "75%"],
            center=["75%", "50%"],
            rosetype="area",
        )
        .set_global_opts(title_opts=opts.TitleOpts(title="Pie-玫瑰图示例"))
    )
    return c


def grid_mutil_yaxis() -> Grid:
    x_data = ["{}月".format(i) for i in range(1, 13)]
    bar = (
        Bar()
        .add_xaxis(x_data)
        .add_yaxis(
            "蒸发量",
            [2.0, 4.9, 7.0, 23.2, 25.6, 76.7, 135.6, 162.2, 32.6, 20.0, 6.4, 3.3],
            yaxis_index=0,
            color="#d14a61",
        )
        .add_yaxis(
            "降水量",
            [2.6, 5.9, 9.0, 26.4, 28.7, 70.7, 175.6, 182.2, 48.7, 18.8, 6.0, 2.3],
            yaxis_index=1,
            color="#5793f3",
        )
        .extend_axis(
            yaxis=opts.AxisOpts(
                name="蒸发量",
                type_="value",
                min_=0,
                max_=250,
                position="right",
                axisline_opts=opts.AxisLineOpts(
                    linestyle_opts=opts.LineStyleOpts(color="#d14a61")
                ),
                axislabel_opts=opts.LabelOpts(formatter="{value} ml"),
            )
        )
        .extend_axis(
            yaxis=opts.AxisOpts(
                type_="value",
                name="温度",
                min_=0,
                max_=25,
                position="left",
                axisline_opts=opts.AxisLineOpts(
                    linestyle_opts=opts.LineStyleOpts(color="#675bba")
                ),
                axislabel_opts=opts.LabelOpts(formatter="{value} °C"),
                splitline_opts=opts.SplitLineOpts(
                    is_show=True, linestyle_opts=opts.LineStyleOpts(opacity=1)
                ),
            )
        )
        .set_global_opts(
            yaxis_opts=opts.AxisOpts(
                name="降水量",
                min_=0,
                max_=250,
                position="right",
                offset=80,
                axisline_opts=opts.AxisLineOpts(
                    linestyle_opts=opts.LineStyleOpts(color="#5793f3")
                ),
                axislabel_opts=opts.LabelOpts(formatter="{value} ml"),
            ),
            title_opts=opts.TitleOpts(title="Grid-多 Y 轴示例"),
            tooltip_opts=opts.TooltipOpts(trigger="axis", axis_pointer_type="cross"),
        )
    )

    line = (
        Line()
        .add_xaxis(x_data)
        .add_yaxis(
            "平均温度",
            [2.0, 2.2, 3.3, 4.5, 6.3, 10.2, 20.3, 23.4, 23.0, 16.5, 12.0, 6.2],
            yaxis_index=2,
            color="#675bba",
            label_opts=opts.LabelOpts(is_show=False),
        )
    )

    bar.overlap(line)
    return Grid().add(
        bar, opts.GridOpts(pos_left="5%", pos_right="20%"), is_control_axis_index=True
    )


def liquid_data_precision() -> Liquid:
    c = (
        Liquid()
        .add(
            "lq",
            [0.3254],
            label_opts=opts.LabelOpts(
                font_size=50,
                formatter=JsCode(
                    """function (param) {
                        return (Math.floor(param.value * 10000) / 100) + '%';
                    }"""
                ),
                position="inside",
            ),
        )
        .set_global_opts(title_opts=opts.TitleOpts(title="Liquid-数据精度"))
    )
    return c


def table_base() -> Table:
    table = Table()

    headers = ["City name", "Area", "Population", "Annual Rainfall"]
    rows = [
        ["Brisbane", 5905, 1857594, 1146.4],
        ["Adelaide", 1295, 1158259, 600.5],
        ["Darwin", 112, 120900, 1714.7],
        ["Hobart", 1357, 205556, 619.5],
        ["Sydney", 2058, 4336374, 1214.8],
        ["Melbourne", 1566, 3806092, 646.9],
        ["Perth", 5386, 1554769, 869.4],
    ]
    table.add(headers, rows).set_global_opts(
        title_opts=opts.ComponentTitleOpts(title="Table")
    )
    return table


def page_draggable_layout():
    page = Page(layout=Page.SimplePageLayout)
    page.add(
        bar_datazoom_slider(),
        line_markpoint(),
        pie_rosetype(),
        grid_mutil_yaxis(),
        liquid_data_precision(),
        table_base(),
    )
    page.render("page_draggable_layout.html")


if __name__ == "__main__":
    page_draggable_layout()
    
    with open(os.path.join(os.path.abspath("."),"page_draggable_layout.html"),'r+',encoding="utf8") as html:#返回链接的路径，并加入os的路径，同时打开
        html_bf=BeautifulSoup(html,"lxml")#创建 beautifulsoup 对象，Beautiful Soup是python的一个库，最主要的功能是从网页抓取数据
        divs=html_bf.find_all("div")#find_all()找到所有匹配结果出现的地方，找到所有匹配div的地方
        divs[0]["style"]='''width:0px;
                          height:0px;
                          position:absolute;
                          top:70px;left:0px;
                          border-style:solid;
                          border-color:#444444;
                          border-bottom-width:0px;
                          border-top-width:1px;
                          border-left-width:0.5px;
                          border-right-width:0.5px;
                          '''
        divs[1]["style"]='''width:600px;
                          height:350px;
                          position:absolute;
                          top:70px;
                          left:0px;
                          border-style:solid;
                          border-color:#444444;
                          border-bottom-width:0px;
                          border-top-width:1px;
                          border-left-width:0.5px;
                          border-right-width:0.5px;
                          '''
        divs[2]["style"] = "width:600px;height:350px;position:absolute;top:70px;left:600px;border-style:solid;border-color:#444444;border-bottom-width:0px;border-top-width:1px;border-left-width:0.5px;border-right-width:0.5px;"
        divs[3]["style"] = "width:600px;height:350px;position:absolute;top:420px;left:0px;border-style:solid;border-color:#444444;border-bottom-width:0px;border-top-width:1px;border-left-width:0.5px;border-right-width:0.5px;"
        divs[4]["style"] = "width:600px;height:350px;position:absolute;top:420px;left:600px;border-style:solid;border-color:#444444;border-bottom-width:0px;border-top-width:1px;border-left-width:0.5px;border-right-width:0.5px;"
        divs[5]["style"] = "width:600px;height:300px;position:absolute;top:770px;left:0px;border-style:solid;border-color:#444444;border-bottom-width:0px;border-top-width:0px;border-left-width:0.5px;border-right-width:0.5px;"
        divs[6]["style"] = "width:600px;height:300px;position:absolute;top:770px;left:600px;border-style:solid;border-color:#444444;border-bottom-width:0px;border-top-width:0px;border-left-width:0.5px;border-right-width:0.5px;"
        # divs[7]["style"] = "width:600px;height:300px;position:absolute;top:420px;left:1200px;border-style:solid;border-color:#444444;border-bottom-width:0px;border-top-width:0px;border-left-width:0.5px;border-right-width:0.5px;"
        # divs[8]["style"] = "width:600px;height:300px;position:absolute;top:420px;left:1800px;border-style:solid;border-color:#444444;border-bottom-width:0px;border-top-width:0px;border-left-width:0.5px;border-right-width:0.5px;"
        # divs[9]["style"] = "width:600px;height:300px;position:absolute;top:420px;left:2400px;border-style:solid;border-color:#444444;border-bottom-width:0px;border-top-width:0px;border-left-width:0.5px;border-right-width:0.5px;"
        # divs[10]["style"] = "width:600px;height:300px;position:absolute;top:720px;left:0px;border-style:solid;border-color:#444444;border-bottom-width:1px;border-top-width:0px;border-left-width:0.5px;border-right-width:0.5px;"
        # divs[11]["style"] = "width:600px;height:300px;position:absolute;top:720px;left:600px;border-style:solid;border-color:#444444;border-bottom-width:1px;border-top-width:0px;border-left-width:0.5px;border-right-width:0.5px;"
        # divs[12][ "style"] = "width:600px;height:300px;position:absolute;top:720px;left:1200px;border-style:solid;border-color:#444444;border-bottom-width:1px;border-top-width:0px;border-left-width:0.5px;border-right-width:0.5px;"
        # divs[13]["style"] = "width:600px;height:300px;position:absolute;top:720px;left:1800px;border-style:solid;border-color:#444444;border-bottom-width:1px;border-top-width:0px;border-left-width:0.5px;border-right-width:0.5px;"
        # divs[14]["style"] = "width:600px;height:300px;position:absolute;top:720px;left:2400px;border-style:solid;border-color:#444444;border-bottom-width:1px;border-top-width:0px;border-left-width:0.5px;border-right-width:0.5px;"
        
        body = html_bf.find("body")
        body["style"] = "background-color:#FFFFFF"
        div_title = "<div align=\"center\" style=\"width:1200px;left:0px;\">\n<span style=\"font-size:32px;font face=\'黑体\';color:#000\"><b>昆明供电局安监部看板</b></div>" # 修改页面背景色、添加看板标题，以及标题的宽度等。注：需根据看板整体宽度调整标题的宽度，使标题呈现居中效果。
        body.insert(0, BeautifulSoup(div_title, "lxml").div)
        html_new = str(html_bf)
        html.seek(0, 0)
        html.truncate()
        html.write(html_new)
        html.close()
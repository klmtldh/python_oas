import re
import pandas as pd
from pandas.api.types import CategoricalDtype

import os
import sys

import datetime
import time
import calendar
from dateutil import rrule
from dateutil.relativedelta import relativedelta

import pymysql
from sqlalchemy import create_engine

import oas_func
import oas_n_day
import oas_pandas_analysis
import oas_excel_pandas_mysql_exchange
import oas_pptx
import oas_docx


class OasPlan(object):
    """
    功能：实现计划类表单的自动分析，包括：开始时间、结束时间、执行部门、执行人员类别的分析
    实例化参数：analysis_category,在文件地图中的模块名称；
              analysis_name, 在文件地图中的文件名称；
              start_column, 计划类表单中的开始时间列名称；
              end_column，计划类表单中的完成时间列名称；
    """
    def __init__(self,
                 analysis_category,
                 analysis_name,
                 start_column,
                 end_column,
                 hold_column,
                 start_date='2020-1-1',
                 end_date=datetime.date.today()):

        self.analysis_category = analysis_category
        self.analysis_name = analysis_name
        # 数据库备份模块名称
        self.analysis_name_backup = analysis_name + '（数据库备份）'
        # Word、PPT文档模块名称
        self.analysis_name_docx = analysis_name + '执行情况docx'
        self.analysis_name_pptx = analysis_name + '执行情况pptx'
        # 开始时间列名
        self.start_column = start_column
        # 结束时间列名
        self.end_column = end_column
        # 保留字段
        self.hold_column = hold_column
        # 数据开始时间
        self.start_date = start_date
        # 数据结束时间
        self.end_date = end_date
        # 定义今日中文时间类
        self.now_day = oas_n_day.NDay(datetime.date.today())
        # 获取Excel和MySQL相关数据
        # 输入Excel地址和Sheet表计MySQL表名称
        self.excel_name_in, self.sheet_name_in, self.mysql_db_in, self.table_name_in = oas_func.read_file_map(
            self.analysis_category, self.analysis_name, '输入')
        # 备份MySQL表名称及存储Excel地址和Sheet表名称
        self.excel_name_backup, self.sheet_name_backup, self.mysql_db_backup, self.table_name_backup = oas_func.read_file_map(
            self.analysis_category, self.analysis_name_backup, '输出')
        # 分析输入MySQL表名称，及存储docx的名称和地址
        self.excel_name_docx, self.sheet_name_docx, self.mysql_db_docx, self.table_name_docx = oas_func.read_file_map(
            self.analysis_category, self.analysis_name_docx, '输出')
        # 分析输入MySQL表名称，及存储pptx的名称和地址
        self.excel_name_pptx, self.sheet_name_pptx, self.mysql_db_pptx, self.table_name_pptx = oas_func.read_file_map(
            self.analysis_category, self.analysis_name_pptx, '输出')
        # 实例化过程，更新MySQL数据库
        self.excel_update_mysql()
        # 实例化过程，备份MySQL数据库
        self.mysql_backup_excel()
        # 生成分析用sql语句
        self.sql = f'SELECT * FROM {self.table_name_docx}'
        # 从数据库获取分析用数据，并存入df
        dbm = oas_excel_pandas_mysql_exchange.DataFrameBothMysql()
        df = dbm.select_mysql(self.sql)
        # 实例化分析df的类，为后续提供分析
        self.pa = oas_pandas_analysis.PandasAnalysis(df)
        # 数据清洗
        self.df = self.pa.cleaning(blank='both')
        # 缺失值分析
        self.df_null, self.df_null_text = self.analysis_null()
        # 生成执行中和到期的df
        self.df_doing, self.df_end = self.select_doing_end(self.start_date)
        self.df_doing_end_list = [self.df,self.df_doing,self.df_end]
        self.copyright = '作者：孔令'
        self.personal_statement = '人生苦短，我用Python！'

    # 用Excel文档更新Mysql数据
    def excel_update_mysql(self, excel_name=None, sheet_name=None, table_name=None):
        if not all([excel_name, sheet_name, table_name]):
            ebm = oas_excel_pandas_mysql_exchange.ExcelBothMysql(self.excel_name_in,
                                                                 self.sheet_name_in,
                                                                 self.table_name_in)
            ebm.excel_to_mysql()
            print(f'{self.excel_name_in}中{self.sheet_name_in}表更新到{self.mysql_db_in}数据库{self.table_name_in}成功。')

    # 将MySQL数据库中的数据备份到本地
    def mysql_backup_excel(self, table_name=None, save_excel_name=None, save_sheet_name=None):
        if not all([table_name, save_sheet_name, save_sheet_name]):
            ebm = oas_excel_pandas_mysql_exchange.ExcelBothMysql(mysql_table=self.table_name_backup,
                                                                 save_path=self.excel_name_backup,
                                                                 save_sheet_name=self.sheet_name_backup)
            ebm.mysql_to_excel()
        print(f'{self.mysql_db_backup}数据库中{self.table_name_backup}表更新到'
              f'{self.excel_name_backup}Excel文件{self.sheet_name_backup}表成功。')

    def analysis_null(self):
        """
        分析self.df中null值的情况，并转换为DataFrame，
        为后续绘制Charm做准备。
        """
        df_null, df_null_text = self.pa.null_item(self.df, flag=0)
        df_null = pd.DataFrame(df_null).reset_index()
        return df_null, df_null_text

    def analysis_no_arrange(self):
        """
        依据start_column和end_column为Null判断没有安排事项。
        """
        df_no_arrange = self.df[(self.df[self.start_column].isnull()) | (self.df[self.end_column].isnull())]
        df_no_arrange = df_no_arrange.fillna('未填写')
        df_no_arrange_text = self.pa.df_iter(df_no_arrange)
        return df_no_arrange_text

    def select_doing_end(self, s_date, e_date=datetime.date.today()):
        """
        功能：获取执行中，和过期的DataFrame
        :param s_date: Type date start date
        :param e_date: Type date end date
        :return: df_doing,df_end Type DataFrame
        """
        start_date = pd.to_datetime(s_date).date()
        end_date = pd.to_datetime(e_date).date()
        df_doing = self.df[(self.df[self.start_column] >= start_date)
                           & (self.df[self.start_column] <= end_date)
                           ]
        df_end = self.df[(self.df[self.start_column] >= start_date)
                         & (self.df[self.end_column] <= end_date)
                         ]
        return df_doing, df_end

    def table_rat(self, piovt_index, piovt_columns, piovt_columns_values, piovt_values, all_name='合计', df=None):
        """
        功能：计算DataFrame 数据透视表的占比。
        :param piovt_index:
        :param piovt_columns:
        :param piovt_columns_values:
        :param piovt_values:
        :param all_name:
        :param df:
        :return: df_table Type DataFrame
        """
        if df is None:
            opa = oas_pandas_analysis.PandasAnalysis(self.df)
            df_table = opa.piovt_table_all_item(
                self.df,
                piovt_index=piovt_index,
                piovt_columns=piovt_columns,
                piovt_columns_values=piovt_columns_values,
                piovt_values=piovt_values,
                all_name=all_name
            )
        else:
            opa = oas_pandas_analysis.PandasAnalysis(df)
            df_table = opa.piovt_table_all_item(
                df,
                piovt_index=piovt_index,
                piovt_columns=piovt_columns,
                piovt_columns_values=piovt_columns_values,
                piovt_values=piovt_values,
                all_name=all_name
            )

        # 计算完成率
        for item in piovt_columns_values:
            df_table[item+'比例'] = df_table[item] / df_table['合计']
            # 设置完成率为百分数
            df_table[item+'比例'] = df_table[item+'比例'].apply(lambda x: '%.2f%%' % (x * 100))

        return df_table

    def department_table_section_doing_end_rat(self):
        department_table = [
            self.table_rat(['责任科室'], '是否完成', ['是', '否'], '行动计划', '安监部', self.df),
            self.table_rat(['责任科室'], '是否完成', ['是', '否'], '行动计划', '安监部', self.df_doing),
            self.table_rat(['责任科室'], '是否完成', ['是', '否'], '行动计划', '安监部', self.df_end)
        ]
        return department_table, self.rat_text_item(department_table, '责任科室')

    def person_table_section_doing_end_rat(self):
        person_table = [
            self.table_rat(['责任人员'], '是否完成', ['是', '否'], '行动计划', '安监部', self.df),
            self.table_rat(['责任人员'], '是否完成', ['是', '否'], '行动计划', '安监部', self.df_doing),
            self.table_rat(['责任人员'], '是否完成', ['是', '否'], '行动计划', '安监部', self.df_end)
        ]
        return person_table

    def rat_text_item(self, df_list, columns, total_name='安监部'):
        """
        功能：将df,df_doing,df_end转换成文字。
        :param df_list: [df,df_doing,df_end]
        :param columns: 转换列
        :param total_name: 合计行名字
        :return: Type str
        """
        title_text = ['', '正在执行中', '到期']
        text_list = []
        item_text = ''
        total_text = ''
        for i, item in enumerate(df_list):
            item_text_list = []
            for j, it in enumerate(list(set(item[columns]))):
                if it == total_name:
                    total_text = '{0}{1}共有{2}条事项{3}，截至{4}，完成{5}条，未完成{6}条，完成率{7};其中：'\
                        .format(self.now_day.get_chinese_date()[0],
                                self.analysis_name,
                                item.loc[item[columns] == total_name, '合计'].values[0],
                                title_text[i],
                                self.now_day.get_chinese_date()[3],
                                item.loc[item[columns] == total_name, '是'].values[0],
                                item.loc[item[columns] == total_name, '否'].values[0],
                                item.loc[item[columns] == total_name, '是比例'].values[0],
                                )
                else:
                    item_text = '{0}共有{1}条事项，完成{2}条，未完成{3}条，完成率{4}。' \
                        .format(it,
                                item.loc[item[columns] == it, '合计'].values[0],
                                item.loc[item[columns] == it, '是'].values[0],
                                item.loc[item[columns] == it, '否'].values[0],
                                item.loc[item[columns] == it, '是比例'].values[0],
                                )
                item_text_list.append(item_text)
                item_text_all = total_text+''.join(item_text_list)
            text_list.append(item_text_all)
        return text_list

    def write_docx(self):
        doc = oas_docx.OasDocx()
        doc.hd(self.now_day.get_chinese_date()[0]+self.analysis_name + '执行情况', font_size=22)
        doc.hd('安全监管部（应急指挥中心）', font_size=16)
        doc.hd(self.now_day.get_chinese_date()[3], font_size=16)
        null_text = f'{self.analysis_name}共有{len(self.df)}条数据，以下数据为{self.df_null_text}'
        doc.par('一、数据总体情况', bold=True)
        doc.par(null_text)
        doc.par('二、未安排事项', bold=True)
        if self.analysis_no_arrange:
            doc.par('工作已经全部安排，无未安排事项。')
        else:
            doc.par(self.analysis_no_arrange)
        doc.par('三、周期事项事项', bold=True)
        # 具体表格
        df_doing_end = [self.df, self.df_doing, self.df_end]
        # 部门/科室完成及完成率数据
        df_table_department, df_table_department_text = self.department_table_section_doing_end_rat()
        # 个人完成及完成率数据
        df_table_person = self.person_table_section_doing_end_rat()

        cycle_text_list = [
            ['三、到期完成情况——科室', '三、到期完成情况——个人', '三、到期完成情况——个人未完成项'],
            ['四、执行中完成情况——科室', '四、执行中完成情况——个人', '四、执行中完成情况——个人未完成项'],
            ['五、全部完成情况——科室', '五、全部完成情况——个人', '五、全部完成情况——个人未完成项']
        ]

        for l in range(3):
            doc.par(cycle_text_list[l][0])
            doc.par(df_table_department_text[2 - l])
            doc.d_table(df_table_department[2 - l])
            doc.par(cycle_text_list[l][1])
            # 生成个人柱状图
            # 插入个人到期完成情况
            doc.d_table(df_table_person[2 - l])

            # 最终输出的列
            hold_columns_out = oas_func.list_del_list(self.hold_column, ['责任科室', '责任人员'])
            df_end_out = df_doing_end[2 - l][self.hold_column]
            df_end_out = df_end_out[df_end_out['是否完成'] == '否']
            duty_department_item = list(set(df_end_out['责任科室']))
            for i, item in enumerate(duty_department_item):
                duty_department_name_item = list(set(df_end_out[df_end_out['责任科室'] == item]['责任人员']))
                for j, name in enumerate(duty_department_name_item):
                    df_end_out_s = df_end_out[(df_end_out['责任科室'] == item)
                                              & (df_end_out['责任人员'] == name)]
                    df_end_out_s = df_end_out_s[hold_columns_out]
                    df_end_text = self.pa.df_iter(df_end_out_s, num=False)
                    for k, text in enumerate(df_end_text):
                        doc.par(cycle_text_list[l][2])
                        doc.par('（' + oas_func.number_to_chinese_number(i + 1) + '）' + item)
                        doc.par(str(j + 1) + '.' + name)
                        doc.par('（' + str(k + 1) + '）' + text)

        doc.par('请各位按照时限完成各项工作！')
        doc.par('人生苦短，我用Python')
        doc.par('本文档由Python自动生成！'
                   '追求极致效率，欢迎志同道合者，共同改进！'
                   '作者：孔令'
                   '版本：V0.5'
                   '手机号：13759176595'
                   '邮箱：klmtldh@163.com'
                )

        doc.save_docx(self.excel_name_docx)

    # 将分析后数据写入PPT。
    def write_pptx(self):

        # 建立PPTX模板
        op = oas_pptx.OasPptx(r'D:\JGY\600-Data\004-auxiliary辅助文件\南方电网logo（16，9）模板.pptx')
        # 模板首页文字字符串
        first_text_list = [
            self.now_day.get_chinese_date()[0]+self.analysis_name + '执行情况',
            '安全监管部（应急指挥中心）',
            self.now_day.get_chinese_date()[3]
        ]
        # 用模板首页文字字符串替换首页文字
        op.alter_csg_logo_first_page(first_text_list)
        # 生成空缺值数据文字
        null_text = f'{self.analysis_name}共有{len(self.df)}条数据，以下数据为{self.df_null_text}'
        # 第一页标题
        op.text_n('一、数据总体情况', font_bold=True)
        # 第一页文字
        op.text_n(null_text, textbox_top=12)
        # 第一页图表
        op.chart(self.df_null, chart_class='bar', top=35)
        op.page()

        # 未安排事项
        op.text_n('二、未安排事项', font_bold=True)
        if self.analysis_no_arrange:
            op.text_n('工作已经全部安排，无未安排事项。', textbox_top=12)
        else:
            op.text_n(self.analysis_no_arrange, textbox_top=12)
        op.page()

        # 具体表格
        df_doing_end = [self.df, self.df_doing, self.df_end]
        # 部门/科室完成及完成率数据
        df_table_department, df_table_department_text = self.department_table_section_doing_end_rat()
        # 个人完成及完成率数据
        df_table_person = self.person_table_section_doing_end_rat()

        cycle_text_list = [
            ['三、到期完成情况——科室', '三、到期完成情况——个人', '三、到期完成情况——个人未完成项'],
            ['四、执行中完成情况——科室', '四、执行中完成情况——个人', '四、执行中完成情况——个人未完成项'],
            ['五、全部完成情况——科室', '五、全部完成情况——个人', '五、全部完成情况——个人未完成项']
        ]

        for l in range(3):
            op.text(cycle_text_list[l][0], font_bold=True)
            op.text_n(df_table_department_text[2-l], textbox_top=12)
            op.p_table(df_table_department[2-l], top=50, height=40)
            op.page()

            op.text(cycle_text_list[l][1], font_bold=True)
            # 生成个人柱状图
            op.chart(
                df_table_person[2-l][['责任人员', '是', '否', '合计']].head(len(df_table_person[2]) - 1),
                chart_class='bar',
                top=12,
                height=40
            )
            # 插入个人到期完成情况
            op.p_table(df_table_person[2-l], top=55, height=40)
            op.page()
            # 最终输出的列
            hold_columns_out = oas_func.list_del_list(self.hold_column, ['责任科室', '责任人员'])
            df_end_out = df_doing_end[2-l][self.hold_column]
            df_end_out = df_end_out[df_end_out['是否完成'] == '否']
            duty_department_item = list(set(df_end_out['责任科室']))
            for i, item in enumerate(duty_department_item):
                duty_department_name_item = list(set(df_end_out[df_end_out['责任科室'] == item]['责任人员']))
                for j, name in enumerate(duty_department_name_item):
                    df_end_out_s = df_end_out[(df_end_out['责任科室'] == item)
                                              & (df_end_out['责任人员'] == name)]
                    df_end_out_s = df_end_out_s[hold_columns_out]
                    df_end_text = self.pa.df_iter(df_end_out_s, num=False)
                    for k, text in enumerate(df_end_text):
                        op.text(cycle_text_list[l][2])
                        op.text('（' + oas_func.number_to_chinese_number(i + 1) + '）' + item, textbox_top=12)
                        op.text(str(j + 1) + '.' + name, textbox_top=18)
                        op.text('（'+str(k + 1) + '）' + text, textbox_top=24)
                        op.page()

        op.text('请各位按照时限完成各项工作！', font_bold=True, alig='center', font_size=45, textbox_top=45)
        op.page()
        op.text('人生苦短，我用Python',font_bold=True)
        op.text_n(['本文档由Python自动生成！',
                   '追求极致效率，欢迎志同道合者，共同改进！',
                   '作者：孔令',
                   '版本：V0.5',
                   '手机号：13759176595',
                   '邮箱：klmtldh@163.com'],
                  font_bold=True,
                  alig='center',
                  font_size=45,
                  textbox_top=12)
        op.save_pptx(self.excel_name_pptx)


def main():
    analysis_category = '工作计划'
    analysis_name = '昆明供电局安全监管重点工作计划'
    hold_column = ['关键任务', '行动计划', '开始时间', '完成时间', '是否完成', '责任科室', '责任人员', '配合人员', '目前推进情况']
    plan = OasPlan(analysis_category, analysis_name, '开始时间', '完成时间', hold_column)
    plan.write_docx()
    plan.write_pptx()
    analysis_category = '问题库'
    analysis_name = '昆明供电局安监部立行立改问题库'
    hold_column = ['发现问题', '行动计划', '开始时间', '完成时间', '是否完成', '责任科室', '责任人员', '配合人员', '目前推进情况']
    plan1 = OasPlan(analysis_category, analysis_name, '开始时间', '完成时间', hold_column )
    plan1.write_docx()
    plan1.write_pptx()

if __name__ == '__main__':
    main()


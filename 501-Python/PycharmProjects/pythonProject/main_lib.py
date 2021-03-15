import kl_class
import configparser

import re
import pandas as pd
from pandas.api.types import CategoricalDtype

import os

import datetime
import time
import calendar
from dateutil import rrule
from dateutil.relativedelta import relativedelta

import pymysql
from sqlalchemy import create_engine

import docx  # 用于操作word文档
from docx.shared import Cm, Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn





# 用于获取电脑名称
def get_pc_name():
    return os.environ['COMPUTERNAME']


# 用于获取配置文件路径，区分办公室和个人电脑
def ini_path():
    if get_pc_name() == 'LAPTOP-HF9P6H1P':
        in_pt = r'D:\JGY\600-Data\001-ini配置文件\孔令配置.ini'
    else:
        in_pt = r'Z:\600数据库\000配置文件\昆明局安监部配置文件.ini'
    return in_pt


# 数字转中文数字
def digital_to_chinese(num):
    num_dict = {'1': '一',
                '2': '二',
                '3': '三',
                '4': '四',
                '5': '五',
                '6': '六',
                '7': '七',
                '8': '八',
                '9': '九',
                '0': '〇',
                }
    index_dict = {1: '',
                  2: '十',
                  3: '百',
                  4: '千',
                  5: '万',
                  6: '十',
                  7: '百',
                  8: '千',
                  9: '亿'
                  }
    nums = list(str(num))
    nums_index = [x for x in range(1, len(nums) + 1)][-1::-1]
    str_ = ''
    for index, item in enumerate(nums):
        str_ = "".join((str_, num_dict[item], index_dict[nums_index[index]]))

    str_ = re.sub("〇[十百千〇]*", "〇", str_)
    str_ = re.sub("〇万", "万", str_)
    str_ = re.sub("亿万", "亿〇", str_)
    str_ = re.sub("〇〇", "〇", str_)
    str_ = re.sub("〇\\b", "", str_)
    return str_


# 实现删除list中指定list
def list_del_list(o_list, d_list):
    r_list = []
    for i in o_list:
        if i not in d_list:
            r_list.append(i)
    return r_list


# 获取模块路径，电脑路径，Sheet，MySQL数据库及表格
def read_file_map(module_name, file_name, in_or_out):
    # 模块用时计时
    start_time = time.time()
    # 读取配置文件ini
    config = configparser.ConfigParser(interpolation=configparser.ExtendedInterpolation())
    config.read(ini_path(), encoding='utf-8')
    path = config.get('文件路径', '工作输入文件地图目录')
    name = os.path.join(path, '昆明供电局安监部文件地图.xlsx')
    ef = pd.ExcelFile(name, engine='openpyxl')
    df = ef.parse('记录')
    if get_pc_name() == 'LAPTOP-HF9P6H1P':
        excel_path = ''.join(
            df[
                (df['模块'] == module_name) &
                (df['文件名称'] == file_name) &
                (df['方式'] == in_or_out)
            ]['个人电脑路径'].values)
    else:
        excel_path = ''.join(
            df[
                (df['模块'] == module_name) &
                (df['文件名称'] == file_name) &
                (df['方式'] == in_or_out)
                ]['服务器路径'].values)

    excel_sheet = ''.join(
        df[
            (df['模块'] == module_name) &
            (df['文件名称'] == file_name) &
            (df['方式'] == in_or_out)
            ]['Sheet名称'].values)

    mysql_db = ''.join(
        df[
            (df['模块'] == module_name) &
            (df['文件名称'] == file_name) &
            (df['方式'] == in_or_out)
            ]['MySQL数据库名称'].values)
    mysql_table = ''.join(
        df[
            (df['模块'] == module_name) &
            (df['文件名称'] == file_name) &
            (df['方式'] == in_or_out)
            ]['MySQL表名称'].values)
    # 模块用时
    elapsed_time = "read_file_map模块用时{0}秒".format(time.time()-start_time)
    print(elapsed_time)
    return excel_path, excel_sheet, mysql_db, mysql_table


# 用Excel文档更新Mysql数据
def update_excel(excel_name, sheet_name, table_name):
    # 模块用时计时
    start_time = time.time()
    ebm = ExcelBothMysql(excel_name, sheet_name, table_name)
    ebm.excel_to_mysql()
    # 模块用时
    elapsed_time = "update_excel模块用时{0}秒".format(time.time()-start_time)
    print(elapsed_time)


def backup_mysql(table_name, save_excel_name, save_sheet_name):
    # 模块用时计时
    start_time = time.time()
    ebm = ExcelBothMysql(mysql_table=table_name, save_path=save_excel_name, save_sheet_name=save_sheet_name)
    ebm.mysql_to_excel()
    print('备份成功')
    # 模块用时
    elapsed_time = "backup_mysql模块用时{0}秒".format(time.time() - start_time)

    print(elapsed_time)


def analysis_null(sql):
    # 模块用时计时
    start_time = time.time()
    dbm = DataFrameBothMysql()
    df = dbm.select_mysql(sql)
    columns_list = list(df)
    pa = PandasAnalysis(df)
    df_null, df_null_text = pa.null_item(flag=0)
    #pdf.df_fig(df_null, 'bar')
    # 模块用时
    elapsed_time = "analysis_null模块用时{0}秒".format(time.time()-start_time)
    print(elapsed_time)
    return len(df), columns_list, df_null, df_null_text

def analysis_no_arrange(sql):
    dbm = DataFrameBothMysql()
    df = dbm.select_mysql(sql)
    df_section_no_arrange = df[(df['开始时间'].isnull()) | (df['完成时间'].isnull())]
    df_section_no_arrange = df_section_no_arrange.fillna('未填写')
    pa = PandasAnalysis(df)
    return pa.df_iter(df_section_no_arrange)


def analysis_all(sql, section_columns, s_date, e_date=datetime.date.today()):
    dm = DataFrameBothMysql()
    df = dm.select_mysql(sql)
    pa = PandasAnalysis(df)
    df = pa.cleaning(blank='both')
    start_date = pd.to_datetime(s_date).date()
    end_date = pd.to_datetime(e_date).date()
    df_section_doing = df[(df[section_columns[0]] >= start_date)
                          & (df[section_columns[0]] <= end_date)
                          ]
    df_section_over = df[(df[section_columns[0]] >= start_date)
                         & (df[section_columns[1]] <= end_date)
                         ]
    return df, df_section_doing, df_section_over


def analysis_all_name():
    pass


def analysis_section(sql, section_columns, hold_columns, flag, s_date, e_date=datetime.date.today()):
    dm = DataFrameBothMysql()
    df = dm.select_mysql(sql)
    start_date = pd.to_datetime(s_date).date()
    end_date = pd.to_datetime(e_date).date()
    df_section_end = df[(df[section_columns[0]] >= start_date)
                        & (df[section_columns[0]] <= end_date)
                        & (df[section_columns[2]] == flag)
                    ]
    df_section_end = df_section_end.fillna('未填写')
    df_section_end = df_section_end[hold_columns]
    pa = PandasAnalysis(df_section_end)
    return pa.df_iter(df_section_end, num=False)

def analysis_name(sql, name, section_columns, hold_columns, flag, s_date, e_date=datetime.date.today()):
    dm = DataFrameBothMysql()
    df = dm.select_mysql(sql)
    start_date = pd.to_datetime(s_date).date()
    end_date = pd.to_datetime(e_date).date()
    df_section_end = df[(df[section_columns[0]] >= start_date)
                        & (df[section_columns[0]] <= end_date)
                        & (df[section_columns[2]] == flag)
                        & (df[section_columns[3]] == name)
                       ]
    df_section_end = df_section_end.fillna('未填写')
    df_section_end = df_section_end[hold_columns]
    pa = PandasAnalysis(df_section_end)
    return pa.df_iter(df_section_end, num=False)




# 实现Excel文件更新到MySQL中
class ExcelBothMysql(object):
    # __doc__内容
    """
    Excel与MySQL数据库之间交互。
    """
    def __init__(self,
                 read_path=None,
                 read_sheet_name=None,
                 mysql_table=None,
                 save_path=None,
                 save_sheet_name=None
                 ):
        self.read_path = read_path
        self.read_sheet_name = read_sheet_name
        self.mysql_table = mysql_table
        self.save_path = save_path
        self.save_sheet_name = save_sheet_name

    def read_excel_file(self, read_path=None, read_sheet_name=None):
        if all([read_path, read_sheet_name]):
            if read_path.endswith('.xls'):
                read_excel_file_ef = pd.ExcelFile(read_path)
                read_excel_file_df = read_excel_file_ef.parse(read_sheet_name)

            elif read_path.endswith('.xlsx'):
                read_excel_file_ef = pd.ExcelFile(read_path, engine='openpyxl')
                read_excel_file_df = read_excel_file_ef.parse(read_sheet_name)
            else:
                print('你选择的文件不是excel文件，请选后缀为.xls或.xlsx文件')
        else:
            if self.read_path.endswith('.xls'):
                read_excel_file_ef = pd.ExcelFile(self.read_path)
                read_excel_file_df = read_excel_file_ef.parse(self.read_sheet_name)

            elif self.read_path.endswith('.xlsx'):
                read_excel_file_ef = pd.ExcelFile(self.read_path, engine='openpyxl')
                read_excel_file_df = read_excel_file_ef.parse(self.read_sheet_name)
            else:
                print('你选择的文件不是excel文件，请选后缀为.xls或.xlsx文件')

        return read_excel_file_df

    def save_excel_file(self, save_df, save_path=None, save_sheet_name=None):
        if all([save_path, save_sheet_name]):
            save_df.to_excel(save_path, sheet_name=save_sheet_name, index=False, engine='openpyxl')
        else:
            save_df.to_excel(self.save_path, sheet_name=self.save_sheet_name, index=False, engine='openpyxl')

    def excel_to_mysql(self, read_path=None, read_sheet_name=None, mysql_table=None):
        if all([read_path, read_sheet_name, mysql_table]):
            excel_to_mysql_dbm = DataFrameBothMysql()
            excel_to_mysql_dbm.replace_mysql(self.read_excel_file(read_path, read_sheet_name), mysql_table)
        else:
            excel_to_mysql_dbm = DataFrameBothMysql()
            excel_to_mysql_dbm.replace_mysql(self.read_excel_file(), self.mysql_table)

    def mysql_to_excel(self, sql=None, save_path=None, save_sheet_name=None):
        if all([sql, save_path, save_sheet_name]):
            mysql_to_excel_dbm = DataFrameBothMysql()
            mysql_to_excel_df = mysql_to_excel_dbm.select_mysql(sql)
            self.save_excel_file(mysql_to_excel_df, save_path, save_sheet_name)
        else:
            sql = "select * from {0}".format(self.mysql_table)
            mysql_to_excel_dbm = DataFrameBothMysql()
            mysql_to_excel_df = mysql_to_excel_dbm.select_mysql(sql)
            self.save_excel_file(mysql_to_excel_df)




# 实现DataFrame 更新到 MySQL
class DataFrameBothMysql(object):
    def __init__(self):
        # 读取配置文件ini
        config = configparser.ConfigParser(interpolation=configparser.ExtendedInterpolation())
        config.read(ini_path(), encoding='utf-8')
        self.connect = pymysql.connect(
            host=config.get('MySQL', 'host'),
            port=int(config.get('MySQL', 'port')),
            user=config.get('MySQL', 'user'),
            passwd=config.get('MySQL', 'passwd'),
            db=config.get('MySQL', 'db'),
            charset=config.get('MySQL', 'charset')
        )
        con = "mysql+pymysql://{0}:{1}@{2}:{3}/{4}?chart{5}".format(config.get('MySQL', 'user'),
                                                                    config.get('MySQL', 'passwd'),
                                                                    config.get('MySQL', 'host'),
                                                                    config.get('MySQL', 'port'),
                                                                    config.get('MySQL', 'db'),
                                                                    config.get('MySQL', 'charset')
                                                                    )
        self.engine = create_engine(con)

    # 查询mysql数据
    def select_mysql(self, sql):
        df = pd.read_sql(sql, self.engine)
        return df

    # 删除一条数据
    def delete_mysql(self, sql):
        cn = self.connect.cursor()
        cn.execute(sql)
        self.connect.commit()
        cn.close()
        self.connect.close()

    # 追加或更新
    def replace_mysql(self, df, table_name):
        try:
            # 执行SQL语句
            df.to_sql('temp', self.engine, if_exists='replace', index=False)  # 把新数据写入 temp 临时表
            # 替换数据的语句
            cn = self.connect.cursor()
            args1 = f" REPLACE INTO {table_name} SELECT * FROM temp "
            cn.execute(args1)
            args2 = " DROP Table If Exists temp"  # 把临时表删除
            cn.execute(args2)
            # 提交到数据库执行
            self.connect.commit()
            cn.close()
            self.connect.close()
            print('更新成功')

        except Exception as e:
            print('更新失败')
            print(e)
            # 发生错误时回滚
            self.connect.rollback()
            # 关闭数据库连接
            cn.close()
            self.connect.close()


class PandasAnalysis(object):
    # __doc__内容
    """
    作者：1979令狐冲
    E-mail：klmtldh@163.com

    pandas数据分析类
    类名称：pandasAnalysis
    函数：
        __init__(self,df)
        get_df(self) 返回类型pd.DataFrame
        cleaning(self,df,blank=None) 类返回型pd.DataFrame
        sort_list(self,obj,list_sort,column=None) 返回类型pd.DataFrame或类型pd.Series
        split_record(self,df: pd.DataFrame, sp_cloc: str, sign: str)返回类型pd.DataFrame
        plus_record(self,dfp: pd.DataFrame, pr_cloc: str, sign: str)返回类型pd.DataFrame
        split_period(self,df, start, end, interval)返回类型pd.DataFrame
        plus_period(self,dfp: pd.DataFrame, s_date,e_date,pr_cloc: str, sign: str)返回类型pd.DataFrame

    使用方法：
        1.实例化
        import pandas as pd
        import pandasAnalysis as pa
        df=pd.DataFrame(
                        [
                            ['A','B','C',pd.to_datetime('2020-12-31'),pd.to_datetime('2021-1-8'),'完成情况1'],
                            ['D','E','F',pd.to_datetime('2021-2-26'),pd.to_datetime('2021-3-7'),'完成情况2'],
                            ['C','E','F',pd.to_datetime('2021-1-20'),pd.to_datetime('2022-3-8'),'完成情况3'],
                        ],columns=['编号','责任人','关键任务','开始日期','结束日期','完成情况'])
        df_pa=pa.pandasAnalysis(df)
        2.get_df(self)函数
        获取df值，封装内部数据

    """

    def __init__(self, df, df_sort_list=None):
        self.df = df
        self.df_sort_list = df_sort_list
        self.df_cleaning = None


    # 数据清洗，主要删除重复值，删除缺失值，删除空格
    def cleaning(self, df=None, blank=None):
        if df is None:
            df = self.df
        else:
            pass
        # 删除重复值
        df.drop_duplicates(inplace=True)
        # 删除缺失值
        df.dropna(axis=0, how='all', inplace=True)
        # 删除空格
        if blank == 'both':
            df = df.applymap((lambda x: str.strip(x) if isinstance(x, str) else x))
        elif blank == 'all':
            df = df.applymap((lambda x: "".join(x.split()) if isinstance(x, str) else x))
        else:
            pass
        self.df_cleaning=df
        return df

    # 按照指定列及list排序，可以实现Series和DataFrame,也可以实现排序list少数排序列的值。
    @staticmethod
    def sort_list(obj, list_sort, column=None):
        object_ = ''
        if isinstance(obj, pd.Series):
            sort_list_df = pd.DataFrame(obj)
            sort_list_df = sort_list_df.reset_index()
            sort_list_df.columns = ['index', 'values']
            sort_list_df1 = sort_list_df[sort_list_df['index'].isin(list_sort)].copy()
            sort_list_df2 = sort_list_df[~sort_list_df['index'].isin(list_sort)].copy()
            sort_list_df1['index'] = sort_list_df1['index'].astype('category')
            sort_list_df1['index'].cat.set_categories(list_sort, inplace=True)
            sort_list_df1.sort_values('index', ascending=True, inplace=True)
            sort_list_df = pd.concat([sort_list_df1, sort_list_df2])
            object_ = pd.Series(sort_list_df['values'].values, index=sort_list_df['index'])
        elif isinstance(obj, pd.DataFrame):
            for i in range(len(column)):
                list_sort[i] = list_sort[i] + list(set(obj[column[i]]).difference(set(list_sort[i])))
            for i in range(len(column)):
                cat_order = CategoricalDtype(
                    list_sort[i],
                    ordered=True
                )
                obj[column[i]] = obj[column[i]].astype(cat_order)
            object_ = obj.sort_values(column, axis=0, ascending=[True] * (len(column)))
        return object_

    # 遍历DataFrame，并添加中文序号，形成字符
    def df_iter(self, df, num=True):
        df_text = []
        j = 1
        for row in df.iterrows():
            x = []
            for i in range(len(row[1])):
                if num:
                    x.append(row[1].index[i] + digital_to_chinese(j) + '：' + str(row[1][row[1].index[i]]) + '\n')
                else:
                    x.append(row[1].index[i] + '：' + str(row[1][row[1].index[i]]) + '\n')
            df_text.append(''.join(x))
            j += 1
        return df_text

    def split_record(self, sp_cloc: str, sign: str, df=None):
        if df is None:
            df = self.df
        else:
            pass
        # 记录原始dataframe的长度，并作为后面追加数据的基准
        df.reset_index(inplace=True, drop=True)
        y = len(df)
        # 记录原始dataframe的长度
        len_df = len(df)
        for i in range(len(df)):
            # 将需要分列的列数据按照sign分列出来
            # print(i,df.loc[i, sp_cloc])
            list_str = str.split(df.loc[i, sp_cloc], sign)
            # 依据分列出来数据增加行
            for j in range(len(list_str)):
                y = y + j
                # 复制i行数据
                df.loc[y] = df.loc[i]
                # 将y行数据sp_cloc列数据设置为list_str中数据
                df.loc[y, sp_cloc] = list_str[j]
            y = y + 1
        # 取出新增加数据
        df = df.tail(len(df) - len_df)
        # 重置index
        df = df.reset_index(drop=True)
        return df

    # 将某列以外都相同的行以','进行合并
    def plus_record(self, dfp: pd.DataFrame, pr_cloc: str, sign: str):
        # 获取dfp的列名
        list_cl = list(dfp)
        # 去除需要合并列的列名
        list_cl.remove(pr_cloc)
        # 按照去除合并列名进行排序
        dfp.sort_values(list_cl, inplace=True)
        # 按照排序后行，重新设置index
        dfp = dfp.reset_index(drop=True)
        # 复制dfp，用以处理
        dfpc = dfp.copy()
        # 去除合并列
        dfpc.drop([pr_cloc], axis=1, inplace=True)
        # 进行查重处理
        list_dp = dfpc.duplicated()
        # 查找重复分界点
        x = list_dp[list_dp.isin([False])].index
        # 因为没有找到index插入的方法，将分界点index转为list
        list_x = []
        for q in range(len(x)):
            list_x.append(x[q])
        # 主要用于加入最后一条记录index
        list_x.append(len(dfp))

        # print(list_x)
        # x.append(int64(len(dfp)))
        yn = []
        # 循环获取重复记录段数据
        for i in range(len(list_x) - 1):
            # 判断是否有需要合并项
            if (list_x[i + 1] - list_x[i]) > 1:
                # 若有序号间隔大于1，则进入循环
                for j in range(list_x[i + 1] - list_x[i]):
                    # 取出需要合并数据，形成list
                    yn.append(dfp.loc[list_x[i] + j, pr_cloc])
                # 将list合并成以sign为分隔字符串。
                y = sign.join(yn)
                # 将字符串赋给dfp第一列
                dfp.loc[list_x[i], pr_cloc] = y
                # 删除多余项目
                for k in range(list_x[i + 1] - list_x[i] - 1):
                    dfp.drop(list_x[i] + 1 + k, axis=0, inplace=True)
                # 清空记录list
                yn = []
        # 重置index
        dfp = dfp.reset_index(drop=True)

        return dfp

    def get_n_day(self, date_time, n=1, m=0):
        # this_month_start = datetime.datetime(self.date_time.year, self.date_time.month, 1)
        this_month_nday = datetime.datetime(date_time.year, date_time.month, n)  # +datetime.timedelta(days=n)
        this_month_end = datetime.datetime(date_time.year, date_time.month,
                                           calendar.monthrange(date_time.year, date_time.month)[1])
        # n_month_start=this_month_start +relativedelta(months=m)
        n_month_end = this_month_end + relativedelta(months=m)
        n_month_nday = this_month_nday + relativedelta(months=m)
        return n_month_nday, n_month_end

    def get_current_week(self, date_time, n=1, w=0):
        monday, sunday = date_time, date_time
        one_day = datetime.timedelta(days=1)
        while monday.weekday() != 0:
            monday -= one_day
        while sunday.weekday() != 6:
            sunday += one_day
        # 返回当前的星期一和星期天的日期
        week_n = monday + datetime.timedelta(days=n)
        n_week_end = sunday + relativedelta(weeks=w)
        n_week_nday = week_n + relativedelta(weeks=w)
        return n_week_nday, n_week_end

    def split_period(self, df, start, end, interval):
        # 重置df.index，保证后面编号不覆盖。
        df = df.reset_index(drop=True)
        # 确定有几行
        df_scope = len(df)
        # 遍历所有行
        for i in range(df_scope):
            # 判断时间行是时间类型
            if isinstance(df[start][i], datetime.datetime) and isinstance(df[end][i], datetime.datetime):
                # 将开始时间行赋值给变量d_start;结束时间赋值给d_end。
                d_start = df[start][i]
                d_end = df[end][i]
                # 判断间隔值是m——月；w——月；d——日
                if interval == 'm':
                    delta = rrule.rrule(rrule.MONTHLY, dtstart=d_start, until=d_end).count()
                    # 解决跨月问题只要月份不同就判定为跨月
                    if self.get_n_day(df[start][i], m=delta - 1)[1] < df[end][i]:
                        delta = delta + 1
                    loop_delta = 0
                    if delta > 1:
                        loop_delta = delta
                        for j in range(loop_delta):
                            df_scope = j + df_scope
                            df.loc[df_scope] = df.loc[i]
                            # this_month_start,this_month_end = get_n_day(df[start][df_scope])
                            if j == 0:
                                df.loc[df_scope, end] = self.get_n_day(df[start][df_scope])[1]
                            elif j == loop_delta - 1:
                                df.loc[df_scope, start] = self.get_n_day(df[start][df_scope], m=j)[0]
                            else:
                                df.loc[df_scope, start], df.loc[df_scope, end] = self.get_n_day(df[start][df_scope],
                                                                                                m=j)
                        df.drop(index=[i], inplace=True)
                    df_scope = df_scope + 1
                # 判断间隔值是m——月；w——周；d——日
                if interval == 'w':
                    delta = rrule.rrule(rrule.WEEKLY, dtstart=d_start, until=d_end).count()
                    # 解决跨周问题,关键点是开始日期所在周推delta个周后的周末是否小于end日期
                    if self.get_current_week(df[start][i], w=delta - 1)[1] < df[end][i]:
                        delta = delta + 1
                    else:
                        pass

                    loop_delta = 0
                    if delta > 1:
                        loop_delta = delta
                        for j in range(loop_delta):
                            df_scope = j + df_scope
                            df.loc[df_scope] = df.loc[i]
                            if j == 0:
                                df.loc[df_scope, end] = self.get_current_week(df[start][df_scope])[1]
                            elif j == loop_delta - 1:
                                df.loc[df_scope, start] = self.get_current_week(df[start][df_scope], w=j)[0]


                            else:
                                df.loc[df_scope, start], df.loc[df_scope, end] = self.get_current_week(
                                    df[start][df_scope], w=j)

                        df.drop(index=[i], inplace=True)
                    df_scope = df_scope + 1
                if interval == 'd':
                    delta = rrule.rrule(rrule.DAILY, dtstart=d_start, until=d_end).count()
                    loop_delta = 0
                    if delta > 1:
                        loop_delta = delta
                        for j in range(loop_delta):
                            df_scope = j + df_scope
                            df.loc[df_scope] = df.loc[i]
                            df.loc[df_scope, start] = d_start + relativedelta(days=+j)
                            df.loc[df_scope, end] = d_start + relativedelta(days=+j)
                        df.drop(index=[i], inplace=True)
                    df_scope = df_scope + 1
        df.reset_index(inplace=True, drop=True)
        return df

    # 将某列以外都相同的行,按照时间','进行合并
    def plus_period(self, dfp: pd.DataFrame, s_date, e_date, pr_cloc: str, sign: str):
        # 获取dfp的列名
        list_cl = list(dfp)
        # 去除需要合并列的列名
        list_cl.remove(s_date)
        list_cl.remove(e_date)
        list_cl.remove(pr_cloc)
        # 按照去除合并列名进行排序
        dfp.sort_values(list_cl, inplace=True)
        # 按照排序后行，重新设置index
        dfp = dfp.reset_index(drop=True)
        # 复制dfp，用以处理
        dfpc = dfp.copy()
        # 去除合并列
        dfpc.drop([s_date, e_date, pr_cloc], axis=1, inplace=True)
        # 进行查重处理
        list_dp = dfpc.duplicated()
        # 查找重复分界点
        x = list_dp[list_dp.isin([False])].index
        # 因为没有找到index插入的方法，将分界点index转为list
        list_x = []
        for q in range(len(x)):
            list_x.append(x[q])
        # 主要用于加入最后一条记录index
        list_x.append(len(dfp))

        # print(list_x)
        # x.append(int64(len(dfp)))
        yn = []
        # 循环获取重复记录段数据
        for i in range(len(list_x) - 1):
            # 判断是否有需要合并项
            if (list_x[i + 1] - list_x[i]) > 1:
                # 若有序号间隔大于1，则进入循环

                for j in range(list_x[i + 1] - list_x[i]):
                    # 取出需要合并数据，形成list
                    yn.append(dfp.loc[list_x[i] + j, pr_cloc])
                # 将list合并成以sign为分隔字符串。
                y = sign.join(yn)
                # 排序
                dfp_d = dfp.loc[list_x[i]:list_x[i + 1] - 1, :].copy()
                dfp_d.sort_values(s_date, inplace=True)
                dfp_d.reset_index(drop=True, inplace=True)
                # 将字符串赋给dfp第一列
                dfp.loc[list_x[i], pr_cloc] = y
                dfp.loc[list_x[i], s_date] = dfp_d.loc[0, s_date]
                dfp.loc[list_x[i], e_date] = dfp_d.loc[list_x[i + 1] - list_x[i] - 1, e_date]

                # 删除多余项目
                for k in range(list_x[i + 1] - list_x[i] - 1):
                    dfp.drop(list_x[i] + 1 + k, axis=0, inplace=True)
                # 清空记录list
                yn = []
        # 重置index
        dfp = dfp.reset_index(drop=True)

        return dfp

    # 将pandas数据转换为文本
    @staticmethod
    def pandas_text(obj, drop_list=None, index_name=None, unit='项', punctuation=[',', '。']):
        tx = ''
        if isinstance(obj, pd.Series):
            text = ''
            for i in range(len(obj)):
                if i != len(obj) - 1:
                    text += str(obj.index[i]) + ' ' + str(obj[i]) + unit + punctuation[0]
                else:
                    text += str(obj.index[i]) + ' ' + str(obj[i]) + unit + punctuation[1]
            if obj.name is not None:
                text = str(obj.name) + '：' + text
            else:
                pass
            tx = text

        if isinstance(obj, pd.DataFrame):
            text_list = []
            if index_name is None:

                for column in obj.iteritems():
                    if column[0] not in drop_list:
                        text = ''
                        for i in range(len(column[1])):
                            if i != len(column[1]) - 1:
                                text += str(column[1].index[i]) + ' ' + str(column[1][i]) + unit + punctuation[0]
                            else:
                                text += str(column[1].index[i]) + ' ' + str(column[1][i]) + unit + punctuation[1]
                            # print('列名'+column[0],'\n',column[1])
                        text = str(column[1].name) + '：' + text
                        text_list.append(text)
            else:
                obj.set_index(index_name, drop=True, inplace=True)
                for column in obj.iteritems():
                    if column[0] not in drop_list:
                        text = ''
                        for i in range(len(column[1])):
                            if i != len(column[1]) - 1:
                                text += str(column[1].index[i]) + ' ' + str(column[1][i]) + unit + punctuation[0]
                            else:
                                text += str(column[1].index[i]) + ' ' + str(column[1][i]) + unit + punctuation[1]
                            # print('列名'+column[0],'\n',column[1])
                        text = str(column[1].name) + '：' + text
                        text_list.append(text)

            tx = text_list
        return tx

    # 获取DataFrame中null空缺值的个数，返回列表和文字；df为DataFrame，all默认值为1列出全部项，为0时只列出有null值的项。
    def null_item(self, df=None, flag=1):
        if df is None:
            null_ = self.df.isnull().sum()
        else:
            null_ = df.isnull().sum()
        null_.sort_values(ascending=False, inplace=True)
        null_.name = '空缺值'
        if flag:
            null_text = self.pandas_text(null_, '项')
        else:
            null_ = null_[null_.values != 0].copy()
            null_text = self.pandas_text(null_, '项', punctuation=[',', ',']) + '其余数据完整。'
        return null_, null_text






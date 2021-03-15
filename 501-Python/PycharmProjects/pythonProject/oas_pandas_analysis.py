"""
作者：孔令
版本:v0.6
"""
import pandas as pd
from pandas.api.types import CategoricalDtype  # 用于DataFrame排序用。
import numpy as np

import time
from dateutil.relativedelta import relativedelta
import datetime
import calendar
from dateutil import rrule

import matplotlib.pyplot as plt  # 用于pandas 绘图
from pylab import mpl  # 解决中文问题

import oas_func


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
        self.df_cleaning = self.cleaning()

    # 数据清洗，主要删除重复值，删除缺失值，删除空格
    def cleaning(self, df=None, blank=None):
        """
        :param df: DataFrame,if df is None,Use self.df
        :param blank: delete blank all or both
        :return: df:DataFrame
        """
        if df is None:
            df=self.df.copy()
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
        self.df_cleaning = df
        return df

    # 按照指定列及list排序，可以实现Series和DataFrame,也可以实现排序list少数排序列的值。
    @staticmethod
    def sort_list(series_or_dataframe, list_sort, column_sort=None):
        """
        作者：孔令
        版本:v0.6
        功能：按照指定列和排序列表对Series或DataFrame排序;
        参数：series_or_dataframe Type Series or DataFrame,需要排序的数据;
             list_sort 排序列表集合，[[]]形式；
             column 需要排序的列
        返回：series，Type Series或dataframe,Type DataFrame
        """
        # 判断series_or_dataframe是否为pd.Series
        if isinstance(series_or_dataframe, pd.Series):
            # 将pd.Series转换成pd.DataFrame
            sort_list_df = pd.DataFrame(series_or_dataframe)
            # 重置索引，并保留索引
            sort_list_df = sort_list_df.reset_index()
            # 将列名改为index,values
            sort_list_df.columns = ['index', 'values']
            # 选出排序列中数据在排序列表中的数据，为了后续不在其中的不变为NAN
            sort_list_df1 = sort_list_df[sort_list_df['index'].isin(list_sort)].copy()
            # 选出排序列表中数据不在排序列表中的数据
            sort_list_df2 = sort_list_df[~sort_list_df['index'].isin(list_sort)].copy()
            # 将index排序列转换为category数据类型
            sort_list_df1['index'] = sort_list_df1['index'].astype('category')
            # 将index列按照list_sort排序
            sort_list_df1['index'].cat.set_categories(list_sort, inplace=True)
            sort_list_df1.sort_values('index', ascending=True, inplace=True)
            # 合并两个数据
            sort_list_df = pd.concat([sort_list_df1, sort_list_df2])
            # 转回Series类型
            series = pd.Series(sort_list_df['values'].values, index=sort_list_df['index'])
            return series
        # 判断series_or_dataframe是否为pd.DataFrame
        elif isinstance(series_or_dataframe, pd.DataFrame):
            # 将排序列中没有在排序列表的数据加入排序列表
            for i in range(len(column_sort)):
                list_sort[i] = list_sort[i] + list(
                    set(series_or_dataframe[column_sort[i]]).difference(set(list_sort[i])))

            for i in range(len(column_sort)):
                # 将排序列表转换为CategoricalDtype类型
                cat_order = CategoricalDtype(
                    list_sort[i],
                    ordered=True
                )
                # 将排序列转换为排序列表类型
                series_or_dataframe[column_sort[i]] = series_or_dataframe[column_sort[i]].astype(cat_order)
            # 按照排序列表排序
            dataframe = series_or_dataframe.sort_values(column_sort, axis=0, ascending=[True] * (len(column_sort)))
            return dataframe

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

    def piovt_table_all_item(self,df, piovt_index, piovt_columns, piovt_columns_values, piovt_values, all_name=None):
        """
        df:DataFrame，需要透视的表
        piovt_index 透视表需要参与分类的列
        piovt_columns 透视表需要统计的列
        piovt_columns_values 透视表需要统计的列中元素取值
        piovt_values 透视表需要计算的列
        all_name  统计合计，且合计All更改名字
        """

        if all_name:
            # piovt_columns中包括的值
            piovt_columns_items = list(set(df.loc[:, piovt_columns].to_list()))
            if len(piovt_columns_items) != len(piovt_columns_values):
                piovt_columns_diff = list(set(piovt_columns_values).difference(set(piovt_columns_items)))
                for item in piovt_columns_diff:
                    df.loc[len(df)] = '过程测试'
                    df.loc[len(df) - 1, piovt_columns] = item

                # 建立基础透视表
                df_table = df.pivot_table(index=piovt_index,
                                          columns=piovt_columns,
                                          values=piovt_values,
                                          aggfunc=len,
                                          fill_value=0
                                          )
                df_table.drop('过程测试', axis=0, inplace=True)
            else:
                # 建立基础透视表
                df_table = df.pivot_table(index=piovt_index,
                                          columns=piovt_columns,
                                          values=piovt_values,
                                          aggfunc=len,
                                          fill_value=0
                                          )
            # 将index变为列
            df_table.reset_index(inplace=True)
            # 标准透视表列名
            df_table_columns_standard = piovt_index + piovt_columns_values
            # 透视表列名
            df_table = df_table[df_table_columns_standard]
            df_table.columns = piovt_index + piovt_columns_values
            # 对行求和
            df_table.loc[len(df_table)] = df_table[piovt_columns_values].apply(lambda x: x.sum())
            df_table.loc[len(df_table) - 1, piovt_index] = all_name
            # 对列求和
            df_table.loc[:, '合计'] = df_table[piovt_columns_values].apply(lambda x: x.sum(), axis=1)
        else:
            # piovt_columns中包括的值
            piovt_columns_items = list(set(df.loc[:, piovt_columns].to_list()))
            if len(piovt_columns_items) != len(piovt_columns_values):
                piovt_columns_diff = list(set(piovt_columns_values).difference(set(piovt_columns_items)))
                for item in piovt_columns_diff:
                    df.loc[len(df)] = '过程测试'
                    df.loc[len(df) - 1, piovt_columns] = item

                # 建立基础透视表
                df_table = df.pivot_table(index=piovt_index,
                                          columns=piovt_columns,
                                          values=piovt_values,
                                          aggfunc=len,
                                          fill_value=0
                                          )
                df_table.drop('过程测试', axis=0, inplace=True)
            else:
                # 建立基础透视表
                df_table = df.pivot_table(index=piovt_index,
                                          columns=piovt_columns,
                                          values=piovt_values,
                                          aggfunc=len,
                                          fill_value=0
                                          )
            # 将index变为列
            df_table.reset_index(inplace=True)
            print(df_table)
            # 标准透视表列名
            df_table_columns_standard = piovt_index + piovt_columns_values
            # 透视表列名
            df_table = df_table[df_table_columns_standard]
            print(df_table)
            df_table.columns = piovt_index + piovt_columns_values
        return df_table

    # 将pandas数据转换为文本
    def pandas_text(self, obj, drop_list=None, index_name=None, unit='项', punctuation=[',', '。']):
        tx = ''
        if isinstance(obj, pd.Series):
            text = ''
            for i in range(len(obj)):
                if i != len(obj) - 1:
                    text += str(obj.index[i]) + ' ' + str(obj[i]) + unit + punctuation[0]
                else:
                    text += str(obj.index[i]) + ' ' + str(obj[i]) + unit + punctuation[1]
            text = str(obj.name) + '：' + text
            tx = text

        if isinstance(obj, pd.DataFrame):
            text_list = []
            if index_name is None:

                for column in df.iteritems():
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
                df.set_index(index_name, drop=True, inplace=True)
                for column in df.iteritems():
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
    def null_item(self, df, flag=1):
        null_ = df.isnull().sum()
        null_.sort_values(ascending=False, inplace=True)
        null_.name = '空缺值'
        if flag:
            null_text = self.pandas_text(null_, '项')
        else:
            null_ = null_[null_.values != 0].copy()
            null_text = self.pandas_text(null_, '项', punctuation=[',', ',']) + '其余数据完整。'
        return null_, null_text

    # 遍历DataFrame，并添加中文序号，形成字符
    def df_iter(self, df, num=True):
        df_text = []
        j = 1
        for row in df.iterrows():
            x = []
            for i in range(len(row[1])):
                if num:
                    x.append(row[1].index[i] + oas_func.number_to_chinese_number(j) + '：' + str(row[1][row[1].index[i]]) + '\n')
                else:
                    x.append(row[1].index[i] + '：' + str(row[1][row[1].index[i]]) + '\n')
            df_text.append(''.join(x))
            j += 1
        return df_text

    # 生产图片
    def df_fig(self,df,fig_class):
        mpl.rcParams['font.sans-serif'] = ['Microsoft YaHei']
        # 解决图表负号显示不正确问题
        plt.rcParams['axes.unicode_minus'] = False
        ax = df.plot(kind=fig_class,
                     legend=False,
                     figsize=(8,5),
                     rot=20
                     #title='Pie of Weather in London'
                     )
        fig = ax.get_figure()
        if oas_func.get_pc_name() == 'LAPTOP-HF9P6H1P':
            # 个人
            fig.savefig(r'D:\JGY\600-Data\006-temporary临时文件\fig.png')
        else:
            # 单位
            fig.savefig(r'D:\out\fig.png')
        plt.clf()
        return


if __name__ == '__main__':
    print(PandasAnalysis.__doc__)
    df = pd.DataFrame(
        [
            ['A', '孔令、刘 媛媛 ', 'C', pd.to_datetime('2020-12-31'), pd.to_datetime('2021-1-8'), '完成情况1'],
            ['D', '李 黎、李进 、昆 明', 'F', pd.to_datetime('2021-2-26'), pd.to_datetime('2021-3-7'), '完成情况2'],
            ['C', '王 玺', np.nan, pd.to_datetime('2021-1-20'), pd.to_datetime('2022-3-8'), '完成情况3'],
        ], columns=['编号', '责任人', '关键任务', '开始日期', '结束日期', '完成情况'])

    pa = PandasAnalysis(df)
    pa.cleaning(blank='all')
    print(pa.df)
    print(pa.df_cleaning)
    # df2 = pa.cleaning(pa.get_df(), 'both')
    # df3 = pa.cleaning(pa.get_df(), 'all')
    # print(df)
    # print(pa.get_df())
    # print(df1)
    # print(df2)
    # print(df3)
    # ls = ['D', 'A', 'B', 'C']
    # df4 = pa.sort_list(df3, ls, '编号')
    # print(df4)
    # df5 = pa.split_record( '责任人', '、',df4)
    # print(df5)
    # df6 = pa.plus_record(df5, '责任人', '，')
    # print(df6)
    # df7 = pd.DataFrame(
    #     [
    #         ['A', 'B', 'C', pd.to_datetime('2020-12-31'), pd.to_datetime('2021-1-8'), '完成情况1'],
    #         ['D', 'E', 'F', pd.to_datetime('2021-2-26'), pd.to_datetime('2021-3-7'), '完成情况2'],
    #         ['C', 'E', 'F', pd.to_datetime('2021-1-20'), pd.to_datetime('2022-3-8'), '完成情况3'],
    #     ]
    #     , columns=['编号', '责任人', '关键任务', '开始日期', '结束日期', '完成情况'])
    # print(df7)
    # df8 = pa.split_period(df7, '开始日期', '结束日期', 'm')
    # print(df8)
    # df9 = pa.plus_period(df8, '开始日期', '结束日期', '完成情况', '；')
    # print(df9)
    # path = r"D:\JGY\300-Work工作\320-PM项目管理\电网年度运行方式\2021年昆明电网八大运行风险防控措施分解表（审查稿).xlsx"
    # df_risk = pd.read_excel(path, '一、主保护、稳控装置拒动导致电网稳定破坏风险', engine="openpyxl")
    # print(df_risk.head)

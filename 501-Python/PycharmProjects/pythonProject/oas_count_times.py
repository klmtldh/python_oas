import pandas as pd
import oas_pandas_analysis


class CountTimes(object):
    """
    依据给定的计划DataFrame和人员DataFrame，生产到位次数DataFrame
    """

    def __init__(self,
                 df_record, 
                 record_hold_columns, 
                 record_count_columns, 
                 df_count,
                 count_column,
                 count_hold_columns,
                 save_path=r'D:\JGY\600-Data\003-out输出文件\02-work工作\01-system工作系统数据\001-安全监督计划\昆明供电局安全监督人员到位统计.xlsx'
                 ):
        """
        
        :param df_record: 需要统计表
        :param record_hold_columns: 统计表需要保留的列
        :param df_count: 统计项目表
        :param count_hold_columns: 统计项目需要保留列
        :param record_count_columns: 含有统计项目内容的列
        """
        # 将null值填充为空字符串,不然后续筛选将出错。
        df_record[record_count_columns] = df_record[record_count_columns].fillna('')
        self.df_record = df_record
        self.record_hold_columns = record_hold_columns
        self.record_count_columns = record_count_columns
        # 将人员表姓名列中的null值删除
        df_count = df_count.dropna(subset=[count_column])
        # 将人员表姓名转换成list为后续判断做准备
        self.df_count_list = df_count[count_column].tolist()
        # 将人员表保留需要信息
        df_count = df_count[count_hold_columns]
        self.df_count = df_count
        self.count_column = count_column
        self.count_hold_columns = count_hold_columns
        self.save_path = save_path

    def get_count_df(self, item_list=None):

        if not item_list:
            item_list = self.df_count_list
        # 将合并的列分开，直接判断是否在列表中，解决contain不能解决孔令，孔令匀的问题
        df_list = []
        for item in self.record_count_columns:
            rccs = self.df_record[item].str.split(',', expand=True)
            for j in range(rccs.shape[1]):
                df_record1 = self.df_record[rccs[j].isin(item_list)]
                df_list.append(df_record1)
        df_record_count = pd.concat(df_list)
        df_record_count.drop_duplicates(inplace=True)
        return df_record_count

    def get_times(self):
        # 设置记录字典
        dt_count = {}
        # 对安监人员进行循环
        for name in self.df_count_list:
            # 每个循环清空df
            df_n_name = pd.DataFrame()
            # 筛选出含有name的记录筛选出含有name所在单位的记录
            df_n_name = self.get_count_df(item_list=[name])
            # 确保相同姓名人员用单位来区分，还可以进一步用部门、班组区分
            df_n_name = df_n_name[
                df_n_name[r'检查部门/单位'].str.contains(
                    ''.join(self.df_count[self.df_count['姓名'] == name]['单位'].values)
                )]

            # 记录个数
            dt_count[name] = len(df_n_name)

        # 将获得的字典统计值转换成df类型
        count_times = pd.DataFrame.from_dict(dt_count, orient='index')
        # 将名字列从index回归到column
        count_times = count_times.reset_index()
        # 将列重新命名
        count_times.columns = ['姓名', '到位次数']
        # 和人员表进行merge
        count_times = pd.merge(self.df_count, count_times, on='姓名').sort_values(by='到位次数', ascending=False)
        # 对index进行重置
        count_times = count_times.reset_index(drop=True)
        return count_times

    def count_times_table(self):
        pass

    def save_to_excel(self, df):
        df.to_excel(
                self.save_path,
                index=False,
                engine='openpyxl'
        )


if __name__ == '__main__':
    path = r'D:\JGY\600-Data\002-in输入文件\02-work工作\01-system工作系统数据\001-安全监督计划\监督计划 (2.1-2.28全局导出).xlsx'
    ex = pd.ExcelFile(path, engine='openpyxl')
    record = ex.parse('Sheet1')
    path1 = r'D:\JGY\600-Data\002-in输入文件\02-work工作\03-document工作文档\昆明供电局领导及安监人员.xlsx'
    ex_count = pd.ExcelFile(path1)
    count = ex_count.parse('Sheet1', engine='openpyxl')
    ct = CountTimes(record,
                    [r'检查部门/单位', '检查负责人', '检查人', '监督计划状态'],
                    ['检查负责人', '检查人'],
                    count,
                    '姓名',
                    ['单位', '姓名', '岗位', '部门'])
    x = ct.get_count_df()
    y = ct.get_times()
    y = y[y['部门'] == '领导']
    z = y.pivot_table(index=['单位'], values=['到位次数'], aggfunc=['sum'], fill_value=0, margins=True)
    z.to_excel(
        r'D:\JGY\600-Data\002-in输入文件\02-work工作\01-system工作系统数据\001-安全监督计划\监督计划筛选.xlsx',
       
        engine='openpyxl'
    )
    # y.to_excel(
    #     r'D:\JGY\600-Data\002-in输入文件\02-work工作\01-system工作系统数据\001-安全监督计划\监督计划筛选.xlsx',
    #     index=False,
    #     engine='openpyxl'
    # )




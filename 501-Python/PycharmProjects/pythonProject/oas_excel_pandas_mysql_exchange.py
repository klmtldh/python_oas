import oas_func
import configparser
import pandas as pd
import pymysql
from sqlalchemy import create_engine




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
        config.read(oas_func.ini_path(), encoding='utf-8')
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
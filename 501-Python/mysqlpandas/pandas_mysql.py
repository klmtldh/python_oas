# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 14:36:53 2020

@author: klmtl
"""

import pandas as pd
import pymysql
from sqlalchemy import create_engine
# MySQL的用户：root, 密码:kongling8167, 端口：3306,数据库：safe_km,字符集：utf8
#engine = create_engine('mysql+mysqlconnector://root:kongling8167@localhost:3306/safe_km')
engine = create_engine('mysql+pymysql://root:kongling8167@localhost:3306/safe_km?charset=utf8')

sql="select * from employee"
df = pd.read_sql_query(sql, engine)
df1=pd.DataFrame([
    ['娃娃','刘',3,'FM',3000]],
    columns=['FIRST_NAME','LAST_NAME', 'AGE', 'SEX', 'INCOME']
    )
#pd.io.sql.to_sql(df1,"人员信息",engine,index=False,if_exists='append')




db = pymysql.connect(host='localhost',
                     port=3306,
                     user='root', 
                     passwd='kongling8167',
                     db='safe_km', 
                     charset='utf8'
                     )
df_data = pd.read_sql(sql , db)


USER_TABLE_NAME = 'employee'
try:
    # 执行SQL语句
    df1.to_sql('temp', engine, if_exists='replace', index=False) # 把新数据写入 temp 临时表
    connection = db.cursor()
    # 替换数据的语句
    args1 = f" REPLACE INTO {USER_TABLE_NAME} SELECT * FROM temp "
    connection.execute(args1)
    args2 = " DROP Table If Exists temp"# 把临时表删除
    connection.execute(args2)
    # 提交到数据库执行
    db.commit()
    
except:
    # 发生错误时回滚
    db.rollback()
    # 关闭数据库连接

connection.close()



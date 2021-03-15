import pandas as pd
import numpy as np
path=r'D:\JGY\600-Data\002-in输入文件\工作\数据库\违章台账 .xlsx'
xl = pd.ExcelFile(path,engine='openpyxl')
df = xl.parse('Sheet1')
#df["人数"].astype(int)
#df = pd.read_excel(path, sheet_name='记录',engine='openpyxl')
print(df.tail(5))
table=pd.pivot_table(df,index=["违章主体","违章单位（班组）","违章等级"],
                     values=["违章原因"],
                     #columns=["单位"],
                     aggfunc={"违章原因":len},
                     fill_value=0,
                     margins=True
                     )
print(table)
table.info()
table.plot()

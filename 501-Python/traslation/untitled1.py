# -*- coding: utf-8 -*-
"""
Created on Wed Dec 16 12:01:37 2020

@author: Kong Ling

E-Mail:klmtldh@163.com

QQ:1170233

Wechat:klmtldh
"""

import pandas as pd
path=r'D:\JGY\600-Data\002-in输入文件\工作\党建\党的理论应知应会知识学习手册.xlsx'
df=pd.read_excel(path)
for i in range(len(df)):
    if i+1<len(df):
        df['答案'][i]=df['问题'][i+1]
df.dropna(inplace=True)
df.to_excel(r'D:\JGY\600-Data\002-in输入文件\工作\党建\党的理论应知应会知识学习手册1.xlsx')
    

# -*- coding: utf-8 -*-
"""
Created on Thu Dec 24 08:14:14 2020

@author: Kong Ling

E-Mail:klmtldh@163.com

QQ:1170233

Wechat:klmtldh
"""

import re
import pandas as pd

path=r'D:\JGY\600-Data\002-in输入文件\个人\招商银行工资卡.txt'
with open(path,'r',encoding='utf-8') as f:
    text=f.read()
    
text_list=re.split('\n',text)

        
df=pd.DataFrame(text_list,columns=['信息'])  
df['账户'] = df['信息'].str.extract('账户(\d{4})', expand = True)
df['年份'] = df['信息'].str.extract('((?:^\d{4})?-\d)', expand = True)
df['日期'] = df['信息'].str.extract('(\d{2}月\d{2}日)', expand = True)
df['时间'] = df['信息'].str.extract('(\d{2}\:\d{2})', expand = True)
df['支付类型'] = df['信息'].str.extract('(扣款|信用卡扣款|银联扣款|支付扣款|转账汇款|入账工资|POS消费|ATM取款|柜台取款|有卡自助消费|支付宝入账|汇款|入账款项)', expand = True)
df['金额'] = df['信息'].str.extract('人民币(\d+\.\d{2})', expand = True)
def fill_nan_upline(x):
    #for i in range(n):
    global y
    if pd.isnull(x):
        x=y    
    else:
        pass
    y=x   
    return x
y=''
df['年份']=df['年份'].apply(fill_nan_upline)
df=df[(df['信息'].str.contains('招商'))&(~df['信息'].str.contains('验证码'))]
#df=df[df['信息'].str.contains('招商')]
df=df.dropna(subset=['账户'])
df['年份'] = df['年份'].str.extract('(^\d{4})')
df['年份'] = df['年份'].fillna('2020')
df['金额']=df['金额'].astype("float")
df.to_excel(r'D:\JGY\600-Data\002-in输入文件\个人\招商银行工资卡.xls')
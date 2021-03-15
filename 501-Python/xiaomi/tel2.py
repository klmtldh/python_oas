# -*- coding: utf-8 -*-
"""
Created on Thu Dec 24 22:28:42 2020

@author: Kong Ling

E-Mail:klmtldh@163.com

QQ:1170233

Wechat:klmtldh
"""

import re
import pandas as pd

path=r'D:\JGY\600-Data\002-in输入文件\个人\小米手机通讯录.vcf'
with open(path,'r',encoding='utf-8') as f:
    text=f.read()
    
pattern = re.compile(r"(BEGIN:VCARD.*?END:VCARD)",re.S)
#result = pattern.findall(text)

text_list=re.split('\n',text)
t=''.join(text_list)
result = pattern.findall(t)
df=pd.DataFrame(result,columns=['信息']) 
df=df[df['信息'].str.contains('孔令')]
#df=df.iloc[3,:]
df['手机'] = df['信息'].str.extract('(?:mobile:|PAGER:|main:)(.*?)[A-Z]', expand = True)
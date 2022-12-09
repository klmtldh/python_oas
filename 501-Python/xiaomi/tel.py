# -*- coding: utf-8 -*-
"""
Created on Thu Dec 24 11:45:47 2020

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
t=t.replace( ' ' , '' )
result = pattern.findall(t)
df=pd.DataFrame(result,columns=['信息']) 
df['姓名'] = df['信息'].str.extract('FN:(.*?)(?:TITLE|ORG|UID)', expand = True)
df['职务'] = df['信息'].str.extract('TITLE:(.*?)[A-Z]', expand = True)        
df['单位'] = df['信息'].str.extract('ORG:(.*?)[A-Z]', expand = True)
df['微博'] = df['信息'].str.extract('URL:(.*?)[A-Z]', expand = True)      
df['UID'] = df['信息'].str.extract('UID:(.*?)[A-Z]', expand = True)
df['生日'] = df['信息'].str.extract('BDAY;VALUE=DATE:(.*?)[A-Z]', expand = True)  
df['家庭住址'] = df['信息'].str.extract('ADR;TYPE=HOME:(.*?)[A-Z]', expand = True)
df['单位地址'] = df['信息'].str.extract('ADR;TYPE=WORK:(.*?)[A-Z]', expand = True)
#df['手机'] = df['信息'].str.extract('(?:mobile:|PAGER:|main:)(.*?)[A-Z]', expand = True)

#df['手机'] = df['手机'].str.extract('(1[3456789]\\d{9})')
df['手机']= df['信息'].str.findall('(?:mobile:|PAGER:|main:)(.*?)[A-Z]')

i=0
for x in df['手机']:
    
    df['手机'][i]=','.join(x)
    i=i+1

#df['手机'] = df['手机'].str.replace( ' ' , '' ) 
df['手机'] = df['手机'].str.findall('(1[3456789]\\d{9})')
i=0
for x in df['手机']:
    
    df['手机'][i]=','.join(x)
    i=i+1
df['家庭座机'] = df['信息'].str.extract('TEL;TYPE=HOME:(.*?)[A-Z]', expand = True) 
df['单位座机'] = df['信息'].str.extract('TEL;TYPE=WORK:(.*?)[A-Z]', expand = True) 
df['个人电子邮箱'] = df['信息'].str.extract('EMAIL;TYPE=HOME:(.*?)[A-Z]', expand = True)
df['工作电子邮箱'] = df['信息'].str.extract('EMAIL;TYPE=WORK:(.*?)[A-Z]', expand = True)
 
df['备注'] = df['信息'].str.extract('NOTE:(.*?)END:VCARD', expand = True)
df.to_excel(r'D:\JGY\600-Data\002-in输入文件\个人\孔令手机通信录.xls')  
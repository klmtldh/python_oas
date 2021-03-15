# -*- coding: utf-8 -*-
"""
Created on Sun Dec 27 14:07:06 2020

@author: Kong Ling

E-Mail:klmtldh@163.com

QQ:1170233

Wechat:klmtldh
"""

import re
import pandas as pd


#import re
def read_vcf(file_name):
    with open(file_name,'r',encoding='utf-8') as f:
        text=f.read()
    text_list=re.split('\n',text)
    text=''.join(text_list)
    text=text.replace( ' ' , '' )
    pattern = re.compile(r"(BEGIN:VCARD.*?END:VCARD)",re.S)
    text_list = pattern.findall(text)
    return text_list

def list_to_str(list_):
    i=0
    for x in list_:
        
        list_[i]=','.join(x)
        i=i+1
    return list_

def vcf_to_xlsx(file_name,save_name):
    columns_name=['姓名',              
              '昵称',
              '职务',
              '单位',              
              'UID',
              '生日',
              '备注',
              '网站',
              '家庭住址',
              '单位地址',
              '手机',
              '总机',
              '单位传真机',
              '家庭传真机',
              '寻呼机',
              '家庭座机',
              '单位座机',
              '个人电子邮箱',
              '工作电子邮箱',              
              'QQ',
              'AIM',
              'MSN',
              '雅虎',
              'SKYPE',
              '谷歌',
              'ICQ',
              'JABBER',
              '姓名拼音'
              
    ]
    df=pd.DataFrame(read_vcf(file_name),columns=['信息']) 
    tag_list=['BEGIN:VCARD',
               'VERSION:',
               'N:',
               'FN:',
               'NICKNAME:',
               'TITLE:',
               'ORG:',               
               'UID:',
               'BDAY;VALUE=DATE:',
               'NOTE:',
               'URL:',
               'ADR;TYPE=HOME:',
               'ADR;TYPE=WORK:',
               'TEL;TYPE=mobile:|TEL;TYPE=PREF,mobile:',               
               'TEL;TYPE=main:',
               'TEL;TYPE=faxWork:',
               'TEL;TYPE=faxHome:',
               'TEL;TYPE=PAGER:',
               'TEL;TYPE=HOME:',
               'TEL;TYPE=WORK:',
               'EMAIL;TYPE=HOME:',
               'EMAIL;TYPE=WORK:',               
               'X-QQ:',
               'X-AIM:',
               'X-MSN:',
               'X-YAHOO:',
               'X-SKYPE-USERNAME:',
               'X-GOOGLE-TALK:',
               'X-ICQ:',
               'X-JABBER:',
               'X-PHONETIC-FIRST-NAME:',
               'X-PHONETIC-MIDDLE-NAME:',
               'X-PHONETIC-LAST-NAME:',
               'END:VCARD'
               ]
    tag_str='|'.join(tag_list)
    for i in range(7):            
        df[columns_name[i]] = df['信息'].str.extract(
            '{0}(.*?)(?:{1})'.format(tag_list[i+3],tag_str),
            expand = True)
    for i in range(19):
        df[columns_name[i+7]] = df['信息'].str.findall(
            '(?:{0})(.*?)(?:{1})'.format(tag_list[i+10],tag_str)
            )
        list_to_str(df[columns_name[i+7]])
        
    df.to_excel(save_name,index=False,engine='openpyxl') 
    
def xlsx_to_vcf(file_name,save_name):
    st='BEGIN:VCARD\nVERSION:3.0\n'
    ed='END:VCARD'
    df=pd.read_excel(file_name,engine='openpyxl')
    with open(save_name,mode='w+',encoding='utf-8') as f:
        for i in range(len(df)):
            f.write(st)
            nm=str(df['姓名'][i])[0]+';'+str(df['姓名'][i])[1:]+';;;'
            f.write('N:'+nm+'\n')
            f.write('FN:'+df['姓名'][i]+'\n')
            f.write('TITLE:'+df[r'职务/岗位名称'][i]+'\n')
            f.write('ORG:'+str(df['岗位所在单位名称'][i])+str(df['所在部门名称'][i])+str(df[r'内设机构（科室、管理室、分部、班组）'][i])+'\n')
            f.write('BDAY;VALUE=DATE:'+df['生日'][i].strftime('%Y-%m-%d')+'\n')
            f.write('TEL;TYPE=mobile:'+str(int(df['手机号'][i]))+'\n')
            f.write('NOTE:'+'身份证号：'+str(df['身份证号'][i])+'\n')
            f.write(ed+'\n\n\n')
            
if __name__ == '__main__':
    file_name=r'D:\JGY\600-Data\002-in输入文件\个人\小米手机通讯录.vcf'
    save_name=r'D:\JGY\600-Data\002-in输入文件\个人\孔令手机通信录.xlsx'
    vcf_to_xlsx(file_name,save_name)
    

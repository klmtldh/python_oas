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
import numpy as np

class miVcfXlsx():
    def __init__(self):
        self.columns_name=['姓名',              
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
              '互联网通话',
              '姓名拼音'             
    ]
        self.tag_list=['BEGIN:VCARD',
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
               'X-SIP:',
               'X-PHONETIC-FIRST-NAME:',
               'X-PHONETIC-MIDDLE-NAME:',
               'X-PHONETIC-LAST-NAME:',               
               'END:VCARD'
               ]
        self.tag_list_=['BEGIN:VCARD',
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
               'TEL;TYPE=mobile:',               
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
               'X-SIP:',
               'X-PHONETIC-FIRST-NAME:',
               'X-PHONETIC-MIDDLE-NAME:',
               'X-PHONETIC-LAST-NAME:',               
               'END:VCARD'
               ]

    def read_vcf(self,file_name):
        with open(file_name,'r',encoding='utf-8') as f:
            text=f.read()
        text_list=re.split('\n',text)
        text=''.join(text_list)
        text=text.replace( ' ' , '' )
        pattern = re.compile(r"(BEGIN:VCARD.*?END:VCARD)",re.S)
        text_list = pattern.findall(text)
        return text_list

    def list_to_str(self,list_):
        i=0
        for x in list_:
            
            list_[i]=','.join(x)
            i=i+1
        return list_

    def vcf_to_xlsx(self,file_name,save_name):
  
        df=pd.DataFrame(self.read_vcf(file_name),columns=['信息']) 

        tag_str='|'.join(self.tag_list)
        for i in range(7):            
            df[self.columns_name[i]] = df['信息'].str.extract(
                '{0}(.*?)(?:{1})'.format(self.tag_list[i+3],tag_str),
                expand = True)
        for i in range(21):
            df[self.columns_name[i+7]] = df['信息'].str.findall(
                '(?:{0})(.*?)(?:{1})'.format(self.tag_list[i+10],tag_str)
                )
            self.list_to_str(df[self.columns_name[i+7]])
        df=df.drop(columns = '信息')
        #df=df.astype(str)
        df.to_excel(save_name,index=False,engine='openpyxl') 
        
    def xlsx_to_vcf(self,file_name,save_name):
        st='BEGIN:VCARD\nVERSION:3.0\n'
        ed='END:VCARD'
        df=pd.read_excel(file_name,engine='openpyxl')
        #df=str(df)
        with open(save_name,mode='w+',encoding='utf-8') as f:
            for i in range(len(df)):
                f.write(st)
                nm=str(df[self.columns_name[0]][i])[0]+';'+str(df[self.columns_name[0]][i])[1:]+';;;'
                f.write(self.tag_list_[2]+nm+'\n')
                for j in range(7):
                    if not pd.isnull(df[self.columns_name[j]][i]) :
                        f.write(self.tag_list_[j+3]+df[self.columns_name[j]][i]+'\n')
                for j in range(21):
                    if not pd.isnull(df[self.columns_name[j+7]][i]) :
                        if str(df[self.columns_name[j+7]][i]).count(',')!=0:
                            for k in range(str(df[self.columns_name[j+7]][i]).count(',')+1):
                                x_list=str(df[self.columns_name[j+7]][i]).split(',')
                                f.write(self.tag_list_[j+10]+x_list[k]+'\n')
                        else:
                            f.write(self.tag_list_[j+10]+str(df[self.columns_name[j+7]][i])+'\n')
                    else:
                        pass
                    
                f.write(ed+'\n\n\n')
                
if __name__ == '__main__':
    file_name=r'D:\JGY\600-Data\002-in输入文件\个人\小米手机通讯录.vcf'
    file_name1=r'D:\JGY\600-Data\002-in输入文件\个人\小米手机通讯录1.vcf'
    save_name=r'D:\JGY\600-Data\002-in输入文件\个人\孔令手机通信录.xlsx'
    mvx=miVcfXlsx()
    mvx.vcf_to_xlsx(file_name,save_name)
    mvx.xlsx_to_vcf(save_name,file_name1)
    

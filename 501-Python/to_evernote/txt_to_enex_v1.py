# -*- coding: utf-8 -*-
"""
Created on Sun Dec  6 21:26:50 2020

@author: Kong Ling

E-Mail:klmtldh@163.com

QQ:1170233

Wechat:klmtldh
"""


import os
import datetime

#获取指定文件夹下的txt文件名和内容
def get_txt_list(path):
    l_name=[]
    l_text=[]
    for root, dirs, files in os.walk(path):
        for name in files:
            l_name.append(name.split(".")[0])            
            with open(os.path.join(root, name),'r', encoding='UTF-8') as f:
                txt=f.read()
                l_text.append(txt)
    return l_name,l_text


#按照印象笔记的导出格式将txt文件    
def evn_txt_n(title,txt,author):
    dt_now=datetime.datetime.now()
    dt=dt_now.strftime("%Y%m%dT%H%M%SZ")
    
    evn_b_l=[]
    evn_s='''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE en-export SYSTEM "http://xml.evernote.com/pub/evernote-export2.dtd">
<en-export export-date="{DT}" application="Evernote/Windows" version="6.x">
        '''.format(DT=dt)
    for i in range(len(title)):
        
        evn_b='''<note><title>{t}</title><content><![CDATA[<?xml version="1.0" encoding="UTF-8"?>
    <!DOCTYPE en-note SYSTEM "http://xml.evernote.com/pub/enml2.dtd">
    
    <en-note><div>{xt}</div></en-note>]]></content><created>{DT}</created><updated>{DT}</updated><note-attributes><author>{a}</author><source>desktop.win</source><source-application>yinxiang.win32</source-application></note-attributes></note>
            '''.format(t=title[i],xt=txt[i],a=author,DT=dt)
        evn_b_l.append(evn_b)
    evn_b=''.join(evn_b_l)
    return evn_s+evn_b+'</en-export>'

#此处改成你自己的txt文件所在的文件夹
path=r'D:\JGY\500-Code\501-Python\to_evernote\txt'
#此处改成你自己的印象笔记导出文件所在的文件夹
save=r'D:\JGY\500-Code\501-Python\to_evernote\enex'

#获取所有的txt文件名字和txt内容。
l_name,l_text=get_txt_list(path)

#用txt文件名作为印象笔记的笔记名字，将1979令狐冲改成你自己的名字
w=evn_txt_n(l_name,l_text,'1979令狐冲')
#将w改成你自己想存成的名字。
with open(save+'//w.enex','w', encoding='UTF-8') as f:
    f.write(w)





# -*- coding: utf-8 -*-
"""
Created on Sun Dec  6 19:51:47 2020

@author: klmtl
"""
import os



def evn_txt(title,txt,author):
    evn_s='''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE en-export SYSTEM "http://xml.evernote.com/pub/evernote-export2.dtd">
<en-export export-date="20201206T125537Z" application="Evernote/Windows" version="6.x">
        '''
    evn_t='''<note><title>{t}</title><content><![CDATA[<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE en-note SYSTEM "http://xml.evernote.com/pub/enml2.dtd">

<en-note><div>{xt}</div></en-note>]]></content><created>20201206T122250Z</created><updated>20201206T125429Z</updated><note-attributes><author>{a}</author><source>desktop.win</source><source-application>yinxiang.win32</source-application></note-attributes></note>
        '''.format(t=title,xt=txt,a=author)

    return evn_s+evn_t+'</en-export>'
    

path=r'D:\JGY\500-Code\501-Python\to_evernote\txt'
save=r'D:\JGY\500-Code\501-Python\to_evernote\enex'
# path=r'D:\JGY\500-Code\501-Python\to_evernote\1.txt'
# save=r'D:\JGY\500-Code\501-Python\to_evernote\2.enex'

for root, dirs, files in os.walk(path):

    for name in files:               
        with open(os.path.join(root, name),'r', encoding='UTF-8') as f:
            txt=f.read()
        
        saven=save+'\\'+name.split(".")[0]+'.enex'
        with open(saven,'w', encoding='UTF-8') as f:
            f.write(evn_txt(name,txt,'1979令狐冲'))




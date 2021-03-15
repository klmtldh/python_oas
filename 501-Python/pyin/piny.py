# -*- coding: utf-8 -*-
"""
Created on Mon Dec 28 10:07:54 2020

@author: Kong Ling

E-Mail:klmtldh@163.com

QQ:1170233

Wechat:klmtldh
"""

import pypinyin
import pandas as pd


# 不带声调的(style=pypinyin.NORMAL)
def pinyin_c(word):
    s = ''
    for i in pypinyin.pinyin(word, style=pypinyin.NORMAL):
        s += ''.join(i).capitalize()+ " "
    return s


# 不带声调的(style=pypinyin.NORMAL)
def pinyin(word):
    s = ''
    for i in pypinyin.pinyin(word, style=pypinyin.NORMAL):
        s += ''.join(i)
    return s

# 带声调的(默认)
def yinjie(word):
    s = ''
    # heteronym=True开启多音字
    for i in pypinyin.pinyin(word, heteronym=True):
        s = s + ''.join(i).capitalize() + " "
    return s


if __name__ == "__main__":
    file_name=r'D:\JGY\600-Data\002-in输入文件\个人\人员信息.xlsx'
    save_name=r'D:\JGY\600-Data\002-in输入文件\个人\人员信息1.xlsx'
    df=pd.read_excel(file_name,engine='openpyxl')
    print(pinyin("孔令"))
    print(pinyin_c("孔令"))
    print(yinjie("孔令"))
    print(yinjie("刘媛媛"))

    for i in range(len(df)):
        df['name_id'][i]=pinyin(df['name'][i])
        df['name_pinyin'][i]=pinyin_c(df['name'][i])
        df['name_pinyin_tone'][i]=yinjie(df['name'][i])
        
    df.to_excel(save_name,index=False,engine='openpyxl') 
    
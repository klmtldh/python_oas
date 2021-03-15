# -*- coding: utf-8 -*-
"""
Created on Sun Dec 20 13:28:13 2020

@author: Kong Ling

E-Mail:klmtldh@163.com

QQ:1170233

Wechat:klmtldh
"""

import pandas as pd
df=pd.DataFrame([
                [1,2,3],
                [3,4,5],
                [3,5,6],
                [3,6,7],
                ],
                #index=[0,1],
                columns=['列1','列2','列3']
                )
dfn=pd.DataFrame({'列1':[1,3,3,3],
                  '列2':[2,4,5,6],
                  '列3':[3,5,6,7],                 
                  },
               
                #index=[0,1],
                
                )

print(df)
print(df.loc[0])
print(df.loc[0,:])
print(df.loc[0:2,'列1'])
print(df.loc[:,'列1':'列3'])
print(df.loc[:,['列1','列3']])
print(df.loc[[0,3],['列1','列3']])
print(df.loc[df['列1']==1,:])

print(df.iloc[0,:])
print(df.iloc[0:2,:])
print(df.iloc[:,0])

df1=df.loc[df['列1']==1,['列2','列1']]
df2=df.loc[df['列1']==1,:]
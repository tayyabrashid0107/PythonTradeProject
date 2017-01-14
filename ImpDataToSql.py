# -*- coding: utf-8 -*-
"""
Created on Tue Dec 13 11:55:44 2016

@author: tayya
"""
import sqlite3
import pandas as pd
from pandas_datareader import data
import csv   
import numpy as np
conn = sqlite3.connect('Shares2.sqlite3') 


f = open('constituents.csv')
csv_f = csv.reader(f)
names=np.array([])

for row in csv_f:
  print (row[0])
  if(row[0]=='Symbol'):
      print(row[0])
  else:
      names=np.append(names, row[0])
      b=list(names)

a = ['YHOO', "GOOG", "TSLA", "MSFT", "MMM" , "PBCT"]

for i in names:
    Datastored=[]
    a=data.DataReader(i, 'yahoo', "01/01/1990", end=None)
    Datastored.append(a)
    conn = sqlite3.connect('Shares2.sqlite3') 
    a.to_sql(name=i, con=conn, if_exists='replace')

df1 = pd.read_sql_query('SELECT* from MSFT ', con=conn)
df1 
c.close()
conn.close()
    

    


#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Dec 22 08:31:38 2021

@author: lucielu

This program transforms Datastream request table to record form long table
"""

import os
import pandas as pd
import numpy as np
import datetime

path='F:/Dropbox/IO-CEL/data and code/CELu codes/Datastream/request tables/'

#path='/Users/yiliulu/Dropbox/IO-CEL/data and code/CELu codes/python'

os.chdir(path)

#variable='Analyst'
#variable='ADR'

def dstorecord(variable):
    
    table=pd.ExcelFile('F:/Dropbox/IO-CEL/data and code/CELu codes/Datastream/Request tables/{0}.xlsm'.format(variable))
    n=len(table.sheet_names)
    
    for i in range(1,n):
    
      xl=pd.read_excel('F:/Dropbox/IO-CEL/data and code/CELu codes/Datastream/Request tables/{0}.xlsm'.format(variable),sheet_name='Sheet{0}'.format(i), converters={'INDICATOR ADR':str,'INDICATOR - ADR':str})
      print(len(xl))
      
      if variable=='ADR':
          xl=xl.assign(ADR=lambda x: 1*(x.iloc[:,1]=='X')+1*(x.iloc[:,2]=='X'))
          xl.rename(columns={'Type':'DSCD'},inplace=True)
          xl=xl[['DSCD','ADR']]
      else:    
          xl=pd.melt(xl,value_name=variable,var_name='date',value_vars=xl.columns.drop(['Name','DSCD']),id_vars='DSCD')
          xl[variable]=pd.to_numeric(xl[variable],errors='coerce')
          xl['date']=xl['date'].astype('int64')
      xl=xl.dropna()
      
      if i==1:
        mergexl=xl
    
    #convert DS string NA to python missing values
      else:
        mergexl=pd.concat([mergexl,xl])

    
    mergexl.to_csv('F:/Dropbox/IO-CEL/data and code/CELu codes/Datastream/Request tables/{0}.csv'.format(variable),index=False) 


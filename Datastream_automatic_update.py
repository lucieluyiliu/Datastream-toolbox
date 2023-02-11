# -*- coding: utf-8 -*-
"""
Created on Tue Feb  7 18:57:35 2023

@author: Lucie Lu: yiliu.lu@mail.mcgill.ca

This is a python script that downloads daily stock return data from Datastream using Excel request table

Previously I calculated stock market variables using Compustat Global, but since some firm-level variables are only
available in Datastream, here I also calculate market measures using Datastream

Downloading higher frequency data from Datastream excel needs to be done in batches.

Here the example is daily return, and querey is run for each year.

ReuqestTable.xslm is a template request table with lists prespecified in column E. 

"""
import win32com.client as win32 
import os
import pandas as pd
import numpy as np
from datetime import datetime

from time import sleep

os.chdir('F:\\Dropbox\\IO-CEL\\data and code\\CELu codes\\Datastream\\Request tables\\')

path="F:\\Dropbox\\IO-CEL\\data and code\\CELu codes\\Datastream\\Request tables\\"

# Open up Excel and make it visible
excel = win32.Dispatch('Excel.Application')
excel.Visible = True

file="RequestTable.xlsm"  #request table template

workbook = excel.Workbooks.Open(path+file)

variable='Price'

frequency='Daily'

start_year="1995"

end_year="2021"


# download daily return by year

for year in range(start_year,end_year):
    
  workbook = excel.Workbooks.Open(path+file)

  RequestTable=workbook.Worksheets("REQUEST_TABLE") 

  RequestTable.Range("B7:B22").Value="Y"  
  
  RequestTable.Range("F7:F22").Value=variable  
  
  RequestTable.Range("I7:I22").Value=frequency  

  RequestTable.Range("G7:G22").Value=start_date

  RequestTable.Range("H7:H22").Value=end_date

   # invoke macro

  workbook.Application.Run("ProcessRequestTable")

  workbook.SaveAs(path+variable+"_daily_{0}.xlsm".format(year))
  
  workbook.Close()
  
  sleep(10) # so that it doesn't crash?
   
  #excel.Application.Quit()


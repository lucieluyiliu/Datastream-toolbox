# -*- coding: utf-8 -*-
"""
Created on Tue Dec 21 17:01:50 2021
Last modified 20230102
@author: lucielu

Convert security information from worldscope list into one single table.
For non-US countries, the input table is FTSE-ACWI-WSuniverse.xlsm
for US, the input table is WSuniverseUS.xlsm

The script applies filters as described in the online appendix, in particular country-specific name filters.
The filers might not be relevant since the universe is based on WorldScope Universe.

"""

import sys
import os
import pandas as pd
import numpy as np
from datetime import datetime

# import xlsxwriter module
import xlsxwriter

#path='/Users/yiliulu/Dropbox/IO-CEL/data and code/CELu codes/Datastream/Constructing Global Universe'
path='F:/Dropbox/IO-CEL/data and code/CELu codes/Datastream/Constructing Global Universe/'

#path='/Users/yiliulu/Dropbox/IO-CEL/data and code/CELu codes/python'

os.chdir(path)

#last available date after 2000-01-01
start=datetime(2000,1,1)

tmp=pd.read_csv('country_list.csv')
ctryISO=tmp['ISO'].tolist()

#name filters
#common name filters
tmp=pd.read_csv('griffinfilters.csv')

namefilters=tmp['keywords'].tolist()

#country-specific name filters
#country specific filters (honestly, I don't think these are still relevant any more but just in case)

ctry_namefilters={}
ctry_namefilters["BR"]=[" PN"," PNA"," PNB"," PNC"," PND"," PNE"," PNF"," PNG"," RCSA"," RCTB"," PNDEAD",
                " PNADEAD"," PNBDEAD"," PNCDEAD"," PNDDEAD"," PNEDEAD"," PNFDEAD"," PNGDEAD"," BDR"]
ctry_namefilters["CO"]=[" PFCL"," PRIVILEGIADAS"," PRVLG"]
ctry_namefilters["GR"]=[" PR"," PB"]
# Hungary filter doesn't seem to be relevant anymore?
#ctry_namefilters["HU"]=["torzsreszveny"]
ctry_namefilters["ID"]=[" FB"," FBDEAD"," RTS"," RIGHTS"]
ctry_namefilters["IN"]=[" XNH"]
ctry_namefilters["IL"]=[" P1"," 1"," 5"]
ctry_namefilters["KR"]=[" 1P"," 2P"," 3P"," 1PB"," 2PB"," 3PB"," 4PB"," 5PB"," 6PB"," 1PFD"," 1PF"," PF2"," 2PF"]
ctry_namefilters["MX"]=[" ACP"," BCP"," C"," L"," CPO"," O"]
ctry_namefilters["MY"]=[" A"," 'A'"," FB", " XCO", " (XCO)", " XCODEAD", " RIGHTS"]
ctry_namefilters["PE"]=[" INVERSION", " INVN", " INV"]
ctry_namefilters["PH"]=[" PDR"]
ctry_namefilters["PT"]=[" R"," 'R'"]
ctry_namefilters["ZA"]=[" N"," CPF"," OPTS"]
ctry_namefilters["SG"]=[" NCPS"," NCPS100", " NRFD", " FB", " FBDEAD"]
ctry_namefilters["TW"]=[" TDR"," 'TDR'"]
ctry_namefilters["TH"]=[" FB"," FBDEAD"]
ctry_namefilters["AU"]=[" RTS"," DEF"," DFD"," DEFF"," PAID", " PRF"]
ctry_namefilters["DE"]=[" GENUSSCHEINE"," GSH"]
ctry_namefilters["BE"]=[" CONV", " VVPR", " STRIP"," AFV"]
ctry_namefilters["CA"]=[" RTS", " SHS"," VTG"," SBVTG"," SUBD"," SR"," SER"," RECPT"," RECEIPT"," EXH",
                " EXCHANGEABLE", " SPLIT"," INC.FD"]
ctry_namefilters["DK"]=[" VXX", " CSE"]
ctry_namefilters["FI"]=[" USE"]
ctry_namefilters["FR"]=[" ADP"," CI"," CIP"," ORA"," ORCI"," OBSA"," OPCSM"," SGP"," SICAV"," FCP"," FCPR"," FCPE",
                " FCPI"," FCPIMT"," OPCVM"]
ctry_namefilters["IT"]=[" RNC"," RIGHTS"," PV"," RP", "RSP"]
ctry_namefilters["NL"]=[" CERT"," CERTS"," STK"]
ctry_namefilters["NZ"]=[" RTS"]
ctry_namefilters["AT"]=[" PC", " GSH"," GENUSSSCHEINE"]
ctry_namefilters["SE"]=[" VXX", " USE"," CONVERTED", " CONV"]
ctry_namefilters["CH"]=[" USE", " CONVERTED", " CONV", " CONVERSION"]
ctry_namefilters["GB"]=[" PAID", " NV"]

ctryFilterKeys=list(ctry_namefilters.keys())

#typefilter

typefilter=['EQ','ADR','GDR']


table=pd.ExcelFile('FTSE-ACWI-WSuniverse.xlsm')

n=len(table.sheet_names)

for i in range (1,n):  #sheets 1-52
    
  #name filter country
  thisISO=ctryISO[i-1]  
    
  xl=pd.read_excel('FTSE-ACWI-WSuniverse.xlsm',sheet_name='Sheet{0}'.format(i))
  
  xl=xl[~xl['DATASTREAM CODE'].isna()]  #remove error entries
 
  print(thisISO +' has {0} securities in WS list \n'.format(len(xl)))
  
  #remove stocks whose last obs is later than 2000-01-01
  xl=xl[xl['DATE/TIME']>start]  #if nan then eliminated
  
  print(thisISO +' has {0} securities after date filter \n'.format(len(xl)))
 
  
  #name filter general
  
  xl=xl[xl.apply(lambda x:not any(ext in str(x['NAME']) for ext in namefilters),axis=1)]
  
  if thisISO in ctryFilterKeys:
   thisCtryfilter=ctry_namefilters[thisISO]
   xl=xl[xl.apply(lambda x:not any(ext in str(x['NAME']) for ext in thisCtryfilter),axis=1)]
   
  print(thisISO +' has {0} securities after name filter \n'.format(len(xl)))
 
  #type filter, primary quote major security
  
  xl=xl[["DATASTREAM CODE","NAME","ISO COUNTRY CODE","CUSIP","ISIN CODE","SEDOL CODE","SIC CODE 1","MAJOR FLAG","QUOTE INDICATOR","STOCK TYPE",\
        "TICKER SYMBOL"]]
  
  xl=xl[(xl['MAJOR FLAG']=='Y')&(xl['QUOTE INDICATOR']=='P')& (xl.apply(lambda x: x['STOCK TYPE'] in typefilter,axis=1))]
  xl=xl.rename(columns={"DATASTREAM CODE":"DSCD","SIC CODE 1":"SIC",\
                   "ISIN CODE":"ISIN","SEDOL CODE": "SEDOL","TICKER SYMBOL": "TICKER",\
                       "ISO COUNTRY CODE": "ISO"})
    
  print(thisISO +' has {0} securities after type filter \n'.format(len(xl)))
 

  if i==1:
    mergexl=xl
  else:
    mergexl=mergexl.append(xl,sort=False)
    

table=pd.ExcelFile('US-WSuniverse.xlsm')

n=len(table.sheet_names)


for i in range (1,n):  #sheets 
    
  xl=pd.read_excel('US-WSuniverse.xlsm',sheet_name='Sheet{0}'.format(i))
  print('US list {0} has {1} securities initially \n'.format(i,len(xl)))

  xl=xl[~xl['DATASTREAM CODE'].isna()]  #remove error entries
 
  #remove stocks whose last obs is later than 2000-01-01
  xl=xl[xl['DATE/TIME']>start]  #if nan then eliminated
  
  print('US list {0} has {1} securities after date filter \n'.format(i,len(xl)))

  #name filter general
  
  xl=xl[xl.apply(lambda x:not any(ext in str(x['NAME']) for ext in namefilters),axis=1)]
  
  print('US list {0} has {1} securities after name filter \n'.format(i,len(xl)))
  #type filter, primary quote major security
  
  xl=xl[["DATASTREAM CODE","NAME","ISO COUNTRY CODE","CUSIP","ISIN CODE","SEDOL CODE","SIC CODE 1","MAJOR FLAG","QUOTE INDICATOR","STOCK TYPE",\
        "TICKER SYMBOL"]]
  
  xl=xl[(xl['MAJOR FLAG']=='Y')&(xl['QUOTE INDICATOR']=='P')& (xl.apply(lambda x: x['STOCK TYPE'] in typefilter,axis=1))]
  xl=xl.rename(columns={"DATASTREAM CODE":"DSCD","SIC CODE 1":"SIC",\
                   "ISIN CODE":"ISIN","SEDOL CODE": "SEDOL","TICKER SYMBOL": "TICKER",\
                       "ISO COUNTRY CODE": "ISO"})
    
  print('US list {0} has {1} securities after type filter \n'.format(i,len(xl)))
  
  mergexl=mergexl.append(xl,sort=False)
    

mergexl.to_csv('WSuniverse.csv',index=False)   

# put list into xlsx spreadsheet 5000 each?

#workbook = xlsxwriter.Workbook('Datastream lists.xlsx')


n=len(mergexl)
nsheets=int(np.ceil(n/5000))
with pd.ExcelWriter('Datastream list.xlsx', engine='openpyxl', mode='a') as writer:  
   for i in range(1,nsheets+1):
    #worksheet = workbook.add_worksheet('list{0}'.format(i))
     xl=mergexl.iloc[5000*(i-1):min(5000*i,n)]
    
     xl.to_excel(writer, sheet_name='list{0}'.format(i))
    
   writer.save()  




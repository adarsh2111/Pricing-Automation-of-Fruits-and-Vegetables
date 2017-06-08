# -*- coding: utf-8 -*-
"""
Created on Sat May 27 15:48:02 2017

@author: user
"""

import os
import xlrd
import pandas as pd
import numpy as np
import openpyxl
import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
date=input("Enter the Date (Enter in dd-mm-yyyy format):")
os.chdir('C:\\Users\\user\\Desktop\\Ninjacart\\%s'%date)
price=pd.ExcelFile("Pricing File - %s.xlsx "%date)
index=pd.read_csv("index%s.csv"%date)
mprice=price.parse('MarketPrice',skiprows=1)
mpriced=pd.DataFrame(mprice)
indexd=pd.DataFrame(index)

emailfile=[]

for i in range(0,indexd['SKUName'].count()):
    for j in range(0,mpriced['SKUName'].count()):
        if(indexd['SKUName'][i]==mpriced['SKUName'][j]):
            emailfile.append([mpriced['SKUName'][j],mpriced['T-1 WS'][j],mpriced['Retail1'][j],mpriced['WS1'][j],mpriced['Retail2'][j],mpriced['WS2'][j],mpriced['Retail3'][j],mpriced['WS3'][i],mpriced['Retail4'][j],mpriced['WS4'][j],indexd['RetailPrice'][i],indexd['WholeSalePrice'][i],indexd['WholeSalePrice'][i]-mpriced['T-1 WS'][j]])
                
emailfile=pd.DataFrame(emailfile)    
emailfile=emailfile.to_csv("EmailFile%s.csv"%date,header=["SKUName","WS-Index Yest","Retail1","WS1","Retail2","WS2","Retail3","WS3","Retail4","WS4","Retail-Index","WS-Index","Diff WSI (Today-Yest)"],index=False)    
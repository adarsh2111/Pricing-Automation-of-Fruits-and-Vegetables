# -*- coding: utf-8 -*-
"""
Created on Wed May 31 18:05:13 2017

@author: user
"""

# -*- coding: utf-8 -*-
"""
Created on Sat May 27 14:44:56 2017

@author: user
"""


# coding: utf-8


#Python 3.5
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
name="Pricing File - "+date
filename="%s.xlsm"% name
price=pd.ExcelFile(filename)
mprice=price.parse('MarketPrice',skiprows=1)
mpriced=pd.DataFrame(mprice)

index=[]
wincrease=[]
rincrease=[]
wdecrease=[]
rti_less_than_grn=[]
wsi_exceed_rti=[]
mix=[]
email_file=[]
for i in range(0,mpriced['SKUName'].count()):
        if(mpriced['ConversionToKgs'][i]==1):
            weightunit="Kg"
        else:
            weightunit="Pcs"
        if(mpriced['T-1 WS'][i]=='-' and mpriced['WS1'][i]==0):
            index.append([date,"15","Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],weightunit,"-","-","8"])
        else:    
            if(mpriced['T-1 WS'][i]=='-' and mpriced['WS1'][i]!=0):
                index.append([date,"15","Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],weightunit,mpriced['Retail1'][i],mpriced['WS1'][i],"8"])
            if (mpriced['SKUClassification'][i]=="Leaves"):
                index.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],weightunit,mpriced['T-1 NC_Retail'][i],mpriced['T-1 WS'][i],8])
                if(mpriced['GRN_Price'][i]!='-'):
                    if(int(mpriced['T-1 NC_Retail'][i])<=int(mpriced['GRN_Price'][i])):
                        rti_less_than_grn.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],mpriced['T-1 NC_Retail'][i],mpriced['T-1 WS'][i],8])
                if(int(mpriced['T-1 WS'][i])>int(mpriced['T-1 NC_Retail'][i])):
                    wsi_exceed_rti.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],mpriced['T-1 NC_Retail'][i],mpriced['T-1 WS'][i],8])
            
            else:
                small=0
                small_ws=0
                t=0
                c=0
                nc=0
                for j in range(1,5):
                    if(mpriced['WS'+str(j)][i]!=0):
                        t=t+1
                        if(mpriced['GRN_Price'][i]!='-'):
                            if(int(mpriced['WS'+str(j)][i])-int(mpriced['GRN_Price'][i])>=5):
                                c=c+1
                            else:
                                nc=nc+1
                           
                r=0
                if(t==0):
                    continue
                elif(t!=0 and t!=1):
                    for j in range(1,5):
                        if((mpriced['WS'+str(j)][i]==mpriced['T-1 WS'][i])and(mpriced['Retail'+str(j)][i]==mpriced['T-1 NC_Retail'][i])):
                            wd=mpriced['WS'+str(j)][i]-mpriced['T-1 WS'][i]
                            rd=mpriced['Retail'+str(j)][i]-mpriced['T-1 NC_Retail'][i]
                            index.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],weightunit,mpriced['Retail'+str(j)][i],mpriced['WS'+str(j)][i],"8"])
                            if(wd>=5):
                                wincrease.append([mpriced['SKUName'][i],mpriced['T-1 WS'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],mpriced['Retail'+str(j)][i],mpriced['WS'+str(j)][i],mpriced['WS'+str(j)][i]-mpriced['T-1 WS'][i]])
                            elif(wd<=-5):
                                wdecrease.append([mpriced['SKUName'][i],mpriced['T-1 WS'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],mpriced['Retail'+str(j)][i],mpriced['WS'+str(j)][i],mpriced['WS'+str(j)][i]-mpriced['T-1 WS'][i]])
                            elif(rd>=5):
                                rincrease.append(mpriced['SKUName'][i])
                            if(mpriced['GRN_Price'][i]!='-'):
                                if(int(mpriced['Retail'+str(j)][i])<=int(mpriced['GRN_Price'][i])):
                                    rti_less_than_grn.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],mpriced['Retail'+str(j)][i],mpriced['WS'+str(j)][i],8])
                            if(int(mpriced['WS'+str(j)][i])>int(mpriced['Retail'+str(j)][i])):
                                wsi_exceed_rti.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],mpriced['Retail'+str(j)][i],mpriced['WS'+str(j)][i],8])
                            r=1
                            break
                elif(t==1):
                    for j in range(1,5):
                        if(mpriced['WS'+str(j)][i]!=0):
                            wd=mpriced['WS'+str(j)][i]-mpriced['T-1 WS'][i]
                            rd=mpriced['Retail'+str(j)][i]-mpriced['T-1 NC_Retail'][i]
                            index.append([date,"15","Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],weightunit,mpriced['Retail'+str(j)][i],mpriced['WS'+str(j)][i],"8"])
                            if(wd>=5):
                                wincrease.append([mpriced['SKUName'][i],mpriced['T-1 WS'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],mpriced['Retail'+str(j)][i],mpriced['WS'+str(j)][i],mpriced['WS'+str(j)][i]-mpriced['T-1 WS'][i]])
                            elif(wd<=-5):
                                wdecrease.append([mpriced['SKUName'][i],mpriced['T-1 WS'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],mpriced['Retail'+str(j)][i],mpriced['WS'+str(j)][i],mpriced['WS'+str(j)][i]-mpriced['T-1 WS'][i]])
                            elif(rd>=5):
                                rincrease.append(mpriced['SKUName'][i])
                            if(mpriced['GRN_Price'][i]!='-'):
                                if(int(mpriced['Retail'+str(j)][i])<=int(mpriced['GRN_Price'][i])):
                                    rti_less_than_grn.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],mpriced['Retail'+str(j)][i],mpriced['WS'+str(j)][i],8])
                            if(int(mpriced['WS'+str(j)][i])>int(mpriced['Retail'+str(j)][i])):
                                wsi_exceed_rti.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],mpriced['Retail'+str(j)][i],mpriced['WS'+str(j)][i],8])
                            r=1
                            break
                if(r==0):
                    if(mpriced['GRN_Price'][i]!='-'):
                        if(t==c):
                            for j in range(1,5):
                                if (mpriced['WS'+str(j)][i]>int(mpriced['GRN_Price'][i])):
                                    small_ws = mpriced['WS'+str(j)][i]
                                    break

                            for j in range(1,5):
                                if(mpriced['WS'+str(j)][i]==0):
                                    continue
                                if(mpriced['WS'+str(j)][i]<=small_ws):
                                    small_ws= mpriced['WS'+str(j)][i]
                                    small = mpriced['Retail'+str(j)][i]
                                
                            
                            if ((int(small_ws)-int(mpriced['GRN_Price'][i]))>=5):
                                wd=small_ws-mpriced['T-1 WS'][i]
                                rd=small-mpriced['T-1 NC_Retail'][i]
                                index.append([date,"15","Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],weightunit,small,small_ws,"8"])
                                if(wd>=5):
                                    wincrease.append([mpriced['SKUName'][i],mpriced['T-1 WS'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],small,small_ws,small_ws-mpriced['T-1 WS'][i]])
                                elif(wd<=-5):
                                    wdecrease.append([mpriced['SKUName'][i],mpriced['T-1 WS'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],small,small_ws,small_ws-mpriced['T-1 WS'][i]])
                                elif(rd>=5):
                                    rincrease.append(mpriced['SKUName'][i])
                                if(mpriced['GRN_Price'][i]!='-'):
                                    if(int(small)<=int(mpriced['GRN_Price'][i])):
                                        rti_less_than_grn.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],small,small_ws,8])
                                if(int(small_ws)>int(small)):
                                    wsi_exceed_rti.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],small,small_ws,8])
                                
                        elif(t==nc):
                            wd=mpriced['WS1'][i]-mpriced['T-1 WS'][i]
                            rd=mpriced['Retail1'][i]-mpriced['T-1 NC_Retail'][i]
                            index.append([date,"15","Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],weightunit,mpriced['Retail1'][i],mpriced['WS1'][i],"8"])
                            if(wd>=5):
                                wincrease.append([mpriced['SKUName'][i],mpriced['T-1 WS'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['WS1'][i]-mpriced['T-1 WS'][i]])
                            elif(wd<=-5):
                                wdecrease.append([mpriced['SKUName'][i],mpriced['T-1 WS'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['WS1'][i]-mpriced['T-1 WS'][i]])
                            elif(rd>=5):
                                rincrease.append(mpriced['SKUName'][i])
                            if(mpriced['GRN_Price'][i]!='-'):
                                if(int(mpriced['Retail1'][i])<=int(mpriced['GRN_Price'][i])):
                                    rti_less_than_grn.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],mpriced['Retail1'][i],mpriced['WS1'][i],8])
                            if(int(mpriced['WS1'][i])>int(mpriced['Retail1'][i])):
                                wsi_exceed_rti.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],mpriced['Retail1'][i],mpriced['WS1'][i],8])
                            
                        elif(c!=0&c<t&c+nc==t):
                            for j in range(1,5):
                                if (mpriced['WS'+str(j)][i]!=0):
                                    min_dev=mpriced['WS'+str(j)][i]-mpriced['T-1 WS'][i]
                                    break
                            for k in range(1,5):
                                if((mpriced['WS'+str(k)][i]!=0)and(mpriced['WS'+str(k)][i]-mpriced['T-1 WS'][i])<=min_dev):
                                    min_dev=mpriced['WS'+str(k)][i]-mpriced['T-1 WS'][i]
                                    small_ws=mpriced['WS'+str(k)][i]
                                    small=mpriced['Retail'+str(k)][i]
                            wd=small_ws-mpriced['T-1 WS'][i]
                            rd=small-mpriced['T-1 NC_Retail'][i]
                            index.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],weightunit,small,small_ws,8])
                            if(wd>=5):
                                wincrease.append([mpriced['SKUName'][i],mpriced['T-1 WS'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],small,small_ws,small_ws-mpriced['T-1 WS'][i]])
                            elif(wd<=-5):
                                wdecrease.append([mpriced['SKUName'][i],mpriced['T-1 WS'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],small,small_ws,small_ws-mpriced['T-1 WS'][i]])
                            elif(rd>=5):
                                rincrease.append(mpriced['SKUName'][i])
                            if(mpriced['GRN_Price'][i]!='-'):
                                if(int(small)<=int(mpriced['GRN_Price'][i])):
                                    rti_less_than_grn.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],small,small_ws,8])
                            if(int(small_ws)>int(small)):
                                wsi_exceed_rti.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],small,small_ws,8])
                            mix.append([mpriced['SKUID'][i],mpriced['SKUName'][i],mpriced['SKUClassification'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],mpriced['T-3 NC_Retail'][i],mpriced['T-2 NC_Retail'],mpriced['T-1 NC_Retail'][i],mpriced['T-3 WS'][i],mpriced['T-2 WS'][i],mpriced['T-1 WS'][i],mpriced['GRN_Price'][i],small,	small_ws,8])    
                    
                        elif (mpriced['WS1'][i]>int(mpriced['GRN_Price'][i])):
                            wd=mpriced['WS1'][i]-mpriced['T-1 WS'][i]
                            rd=mpriced['Retail1'][i]-mpriced['T-1 NC_Retail'][i]
                            index.append([date,"15","Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],weightunit,mpriced['Retail1'][i],mpriced['WS1'][i],"8"])
                            if(wd>=5):
                                wincrease.append([mpriced['SKUName'][i],mpriced['T-1 WS'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['WS1'][i]-mpriced['T-1 WS'][i]])
                            elif(wd<=-5):
                                wdecrease.append([mpriced['SKUName'][i],mpriced['T-1 WS'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['WS1'][i]-mpriced['T-1 WS'][i]])
                            elif(rd>=5):
                                rincrease.append(mpriced['SKUName'][i])
                            if(mpriced['GRN_Price'][i]!='-'):
                                if(int(mpriced['Retail1'][i])<=int(mpriced['GRN_Price'][i])):
                                    rti_less_than_grn.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],mpriced['Retail1'][i],mpriced['WS1'][i],8])
                            if(int(mpriced['WS1'][i])>int(mpriced['Retail1'][i])):
                                wsi_exceed_rti.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],mpriced['Retail1'][i],mpriced['WS1'][i],8])
                            
                        else:
                            wd=mpriced['WS1'][i]-mpriced['T-1 WS'][i]
                            rd=mpriced['Retail1'][i]-mpriced['T-1 NC_Retail'][i]
                            index.append([date,"15","Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],weightunit,mpriced['Retail2'][i],mpriced['WS2'][i],"8"])
                            email_file.append([mpriced['SKUName'][i],mpriced['T-1 WS'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],mpriced['Retail1'][i],mpriced['WS1'][i],wd])
                            if(wd>=5):
                                wincrease.append([mpriced['SKUName'][i],mpriced['T-1 WS'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['WS1'][i]-mpriced['T-1 WS'][i]])
                            elif(wd<=-5):
                                wdecrease.append([mpriced['SKUName'][i],mpriced['T-1 WS'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['WS1'][i]-mpriced['T-1 WS'][i]])
                            elif(rd>=5):
                                rincrease.append(mpriced['SKUName'][i])
                            if(mpriced['GRN_Price'][i]!='-'):
                                if(int(mpriced['Retail2'][i])<=int(mpriced['GRN_Price'][i])):
                                    rti_less_than_grn.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],mpriced['Retail2'][i],mpriced['WS2'][i],8])
                            if(int(mpriced['WS2'][i])>int(mpriced['Retail2'][i])):
                                wsi_exceed_rti.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],mpriced['Retail2'][i],mpriced['WS2'][i],8])
                    else:
                        wd=mpriced['WS1'][i]-mpriced['T-1 WS'][i]
                        rd=mpriced['Retail1'][i]-mpriced['T-1 NC_Retail'][i]
                        index.append([date,"15","Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],weightunit,mpriced['Retail1'][i],mpriced['WS1'][i],"8"])
                        if(wd>=5):
                            wincrease.append([mpriced['SKUName'][i],mpriced['T-1 WS'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['WS1'][i]-mpriced['T-1 WS'][i]])
                        elif(wd<=-5):
                            wdecrease.append([mpriced['SKUName'][i],mpriced['T-1 WS'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['Retail2'][i],mpriced['WS2'][i],mpriced['Retail3'][i],mpriced['WS3'][i],mpriced['Retail4'][i],mpriced['WS4'][i],mpriced['Retail1'][i],mpriced['WS1'][i],mpriced['WS1'][i]-mpriced['T-1 WS'][i]])
                        elif(rd>=5):
                            rincrease.append(mpriced['SKUName'][i])
                        if(mpriced['GRN_Price'][i]!='-'):
                            if(int(mpriced['Retail1'][i])<=int(mpriced['GRN_Price'][i])):
                                rti_less_than_grn.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],weightunit,mpriced['Retail1'][i],mpriced['WS1'][i],8])
                        if(int(mpriced['WS1'][i])>int(mpriced['Retail1'][i])):
                            wsi_exceed_rti.append([date,15,"Wholesale Index",mpriced['SKUID'][i],mpriced['SKUName'][i],mpriced['Retail1'][i],mpriced['WS1'][i],8])

index=pd.DataFrame(index)
index.to_csv('index%s.csv'%date,header=["PurchaseDate","MarketId","Market","SKUId","SKUName","WeightUnit","RetailPrice","WholeSalePrice","PriceSource"],index=False)
rti_less_than_grn=pd.DataFrame(rti_less_than_grn)
wsi_exceed_rti=pd.DataFrame(wsi_exceed_rti)
mix=pd.DataFrame(mix)
wincrease=pd.DataFrame(wincrease)
wdecrease=pd.DataFrame(wdecrease)
rti_less_than_grn.to_csv("RTI_less_than_GRN.csv",header=["PurchaseDate","MarketId","Market","SKUId","SKUName","RetailPrice","WholeSalePrice","PriceSource"],index=False)
if len( wsi_exceed_rti)==0:
    print("NO WSI exceeds RTI")
else:
    wsi_exceed_rti.to_csv("wsi_exceed_rti.csv",header=["PurchaseDate","MarketId","Market","SKUId","SKUName","RetailPrice","WholeSalePrice","PriceSource"],index=False)
if len(mix)==0:
    print("No mix market SKU")
else:
    mix.to_csv("mix.csv",header=["SKUID","SKUName","SKUClassification","Retail1","	WS1","Retail2","WS2","Retail3","WS3","Retail4","WS4","T-3 NC_Retail","T-2 NC_Retail","	T-1 NC_Retail","T-3 WS","T-2 WS","	T-1 WS","	GRN_Price","Retail-Index","	WS-Index","Price Source"],index=False)
wincrease.to_csv("Wholesale Index Increased by 5 or more List.csv",header=["SKUName","WS-Index Yest","Retail1","WS1","Retail2","WS2","Retail3","WS3","Retail4","WS4","Retail-Index","WS-Index","Diff WSI (Today-Yest)"],index=False)
wdecrease.to_csv("Wholesale Index Decreased by 5 or more List.csv",header=["SKUName","WS-Index Yest","Retail1","WS1","Retail2","WS2","Retail3","WS3","Retail4","WS4","Retail-Index","WS-Index","Diff WSI (Today-Yest)"],index=False)
ofile = openpyxl.load_workbook(filename)
index=pd.read_csv('index%s.csv'%date)
index=pd.DataFrame(index)
marketprice=ofile.get_sheet_by_name('MarketPrice')
for row_index in range(2,index["SKUName"].count()):
    marketprice['T%d'%(row_index+1)]=int(index['WholeSalePrice'][row_index])
    marketprice['U%d'%(row_index+1)]=int(index['RetailPrice'][row_index])
ofile.save(filename)

# In[ ]:




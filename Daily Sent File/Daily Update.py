###Setting Dir
import csv
import pandas as pd
from pandas.tseries.offsets import BDay
import os
dir=os.getcwd()
print(dir)
os.chdir('C:\\Users\\LuGao\\Constanter Philanthropy Services\\Constanter IO - General\\05. Operations\\Testing')
###Importing File
import datetime
LastBusDay=(datetime.date.today() -BDay(1)).strftime('%y%m%d')
ReadDay=(datetime.date.today() -BDay(2)).strftime('%y%m%d')
file_name=ReadDay+'_UnitBalance_Constanter.csv'
send_file=LastBusDay+'_UnitBalance_Constanter.csv'
readFile = pd.read_csv(file_name,sep=',')
readFile['trade_date']= (datetime.date.today() -BDay(2)).strftime('%d/%m/%Y')
print(readFile.iloc[:,1])

###Output
readFile.to_csv(LastBusDay+'_UnitBalance_Constanter.csv',index=False)

###Send Email
###win32 module error to correct __init__.py (C:\Users\LuGao\Projects\JPM\venv\Lib\site-packages\win32com)
###update import win32api,sys,os to from win32 import win32api
###import sys, os
import win32com.client as win32
from pathlib import Path

Attachment_Dir=Path.cwd()
attachment_path=str(Attachment_Dir/f"{LastBusDay}_UnitBalance_Constanter.csv")

outlook=win32.Dispatch('outlook.application')
mail=outlook.CreateItem(0)
mail.To='l.gao@constanter.org'
mail.CC='l.gao@constanter.org'
mail.Subject='test'
mail.Body='Dear Client,\n\nPlease kindly check the attached updated file.\n\nKind Regards,\n\nLu Gao'
mail.Attachments.Add(Source=attachment_path)
mail.Send()




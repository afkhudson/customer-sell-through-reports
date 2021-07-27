import pandas as pd
from datetime import date
import openpyxl
from openpyxl import load_workbook
import glob
import os
import xlwings as xw
import time
import win32com.client as win32
from time import sleep


#GETTING TODAYS DATE FOR FILE NAME
today = date.today()
d4 = today.strftime("%Y%m%d")
print(f'todays date: {d4}')
print('--------------------------------------------')
print('--------------------------------------------')

#open master dashboard


#open data dump FILE
datadump = r'MS - BBY - Data Dump - Copy.xlsx'

#*------------  PASTE INDIVIDUAL CUSTOMER DATA INTO DATA DUMP SHEET

#open weekly customer report
print('LASTEST FILES IN FOLDERS:')

#get latest adorama report
adoramaraw_dir = r"Adorama Raw/*"
adorama_list_of_files = glob.glob(adoramaraw_dir)
adoramaraw_latest = max(adorama_list_of_files, key=os.path.getctime)
print(f'adoramaraw_latest created: {time.ctime(os.path.getmtime(adoramaraw_latest))}')

adoramaraw_df = pd.read_excel(open(adoramaraw_latest,'rb'),skiprows=1)

adorama_disty_column = []
for i in range(len(adoramaraw_df)):
    adorama_disty_column.append('Adorama')

adoramaraw = adoramaraw_df

print(f'rows: {len(adoramaraw)} columns: {len(adoramaraw.columns)}')

#get latest  B&H report
bhraw_dir = r"BH Raw/*"
bh_list_of_files = glob.glob(bhraw_dir)
bhraw_latest = max(bh_list_of_files, key=os.path.getctime)
print(f'bhraw_latest created: {time.ctime(os.path.getmtime(bhraw_latest))}')

bhraw_df = pd.read_excel(open(bhraw_latest,'rb'),skiprows=1)

#creating disty column for B&H
bh_disty_column = []
for i in range(len(bhraw_df)):
    bh_disty_column.append('B&H Photo')

bhraw = bhraw_df
print(f'rows: {len(bhraw)} columns: {len(bhraw.columns)}')

#gets latest  MSFT report
msraw_dir = r"MS Raw/*"
ms_list_of_files = glob.glob(msraw_dir)
msraw_latest = max(ms_list_of_files, key=os.path.getctime)
print(f'msraw_latest created: {time.ctime(os.path.getmtime(msraw_latest))}')

msraw_df = pd.read_excel(open(msraw_latest,'rb'),skiprows=2)
msraw = msraw_df
print(f'rows: {len(msraw)} columns: {len(msraw.columns)}')

#get latest SPA report
sparaw_dir = r"SPA Raw/*"
spa_list_of_files = glob.glob(sparaw_dir)
sparaw_latest = max(spa_list_of_files, key=os.path.getctime)
print(f'sparaw_latest created: {time.ctime(os.path.getmtime(sparaw_latest))}')

sparaw_df = pd.read_excel(open(sparaw_latest,'rb'))
sparows = len(sparaw_df.columns)
sparaw = sparaw_df.iloc[:, lambda df: [1,sparows-9,sparows-4,sparows-2]]
sparaw.columns = ['Paste PN','Sell Thru','Inventory','On Order']
print(f'rows: {len(sparaw)} columns: {len(sparaw.columns)}')

#gets latest bby raw data file
bbyraw_dir = r"BBY Raw/*"

bby_list_of_files = glob.glob(bbyraw_dir)
bbyraw_latest = max(bby_list_of_files, key=os.path.getctime)
print(f'bbyraw_latest created: {time.ctime(os.path.getmtime(bbyraw_latest))}')
#latest file into dataframe
bbyraw = pd.read_csv(open(bbyraw_latest,'rb'))
print(f'rows: {len(bbyraw)} columns: {len(bbyraw.columns)}')

print('--------------------------------------------')
print('--------------------------------------------')
input('Make sure that these are the latest files, press any key to continue..')

#opening datadump to paste in
wb_app = xw.App(visible=False)
wb_app.display_alerts=False

wb = xw.Book(datadump)

print('opening datadump')
wsbby = wb.sheets['BBY Raw']
wsspa = wb.sheets['SPA Raw']
wsms = wb.sheets['MS Raw']

#paste bbyraw pd to datadump bby raw tab
wsbby.range('B1').options(index=False).value = bbyraw
print('pasting bbyraw data')

#paste sparaw to datadump bby raw tab
wsspa.range('C1').options(index=False).value = sparaw
print('pasting sparaw data')

#paste MSFT raw data to datadump
wsms.range('D1').options(index=False).value = msraw
print('pasting msraw data')

#wait for pasting to bne done before save and close
for i in range(3):
    sleep(1);
    print(f'waiting {i}')

#saves datadump file and closes
app = xw.apps.active
wb.save()
app.quit()
print('closing datadump')

#*------------  REFRESH DATA DUMP SHEET FOR CLEAN DATA

#setting datadump to reopen
xlapp = win32.DispatchEx('Excel.Application')
#opening data dummp to refresh pivot tables and then close datadump
sleep(2)
xlbook = xlapp.Workbooks.Open("C:/Users/sosa/OneDrive/Coding/automate-dashboard/MS - BBY - Data Dump - Copy.xlsx")
xlapp.DisplayAlerts = False
xlapp.Visible = True
print('opening datadump to refresh')
xlbook.RefreshAll()
print('refreshing pivottables')
xlbook.Save()
xlapp.Quit()



#*------------   MOVING CLEAN DATA TO ONE SHEET
#reading clean data
sleep(2)
msdf = pd.read_excel(open(datadump,'rb'),sheet_name="MS Clean");
bbydf = pd.read_excel(open(datadump,'rb'),sheet_name="BBY Clean");
spadf = pd.read_excel(open(datadump,'rb'),sheet_name="SPA Clean");

#pulling specific columns from clean data
msCleanData = msdf.iloc[:, lambda df: [0,1,2,3,4]]
bbyCleanData = bbydf.iloc[:, lambda df:[0,1,2,3,4,]]



spaCleanData = spadf.iloc[:, lambda df:[0,1,2,3,4]]
#creating spa note column to differentiate SPA from BBY
spa_column = []
for i in range(len(spaCleanData)):
    spa_column.append('SPA')
spaCleanData.insert(5,"Note",spa_column,True)

bhCleanData = bhraw.iloc[:, lambda df:[2,5,4]]
bhCleanData.insert(0, "Distributor", bh_disty_column, True) #inserting disty column for B&H

adoramaCleanData = adoramaraw.iloc[:, lambda df:[2,3,5]]
adoramaCleanData.insert(0, "Distributor", adorama_disty_column, True) #inserts Disty Column

#grouping clean data
accountsdf = [msCleanData, bbyCleanData, spaCleanData, bhCleanData, adoramaCleanData]


#creating excel sheet with todays date
writer = pd.ExcelWriter(f'{d4} sell thru data.xlsx')
print(f'total accounts: {len(accountsdf)}')

#setting number of lines to skip per set of data
lines = []
for i in range(len(accountsdf)):
    lines.append(len(accountsdf[i]))
print(f'lines to skip: {lines}')

#printing data to sheet and skipping lines with existing data
skip = 0
for i in range(len(accountsdf)):
    print(f'pasting acccount #{i+1} to line {skip+1}')
    accountsdf[i].to_excel(writer, header=None, index=False, startrow=(skip))
    skip += lines[i]

#saving excel file
writer.save()
writer.handles = None
xw.Book(f'{d4} sell thru data.xlsx')
del xlapp

#*------------  PASTE CLEAN ONE SHEET INTO DASHBOARD


#*------------  DASHBOARD DATA CLEAN UP

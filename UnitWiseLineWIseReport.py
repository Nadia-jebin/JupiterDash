import pandas as pd
import csv
import base64
import datetime
import io
import time
from time import mktime
import sys
import numpy as np

import matplotlib.pyplot as plt
import plotly
import chart_studio.plotly as py
import seaborn as sns
from plotly import graph_objs as go
import plotly.express as px
import dash
import dash_table
from datetime import date
from dash import Dash, dcc, html
from dash.dependencies import Input, State, Output
import dash_bootstrap_components as dbc
from dash import callback_context

import base64
import os
from urllib.parse import quote as urlquote

from flask import Flask, send_from_directory
import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, State, Output
import dash_bootstrap_components as dbc
from fractions import Fraction
from math import ceil

import matplotlib.pyplot as plt

import requests
import csv
import json
from urllib.request import urlopen

import openpyxl
from openpyxl import Workbook, load_workbook

import math


production_url='http://10.12.13.164:8110/api/BasicQMSData/BirichinaDataForReport?StartDate=2022-04-12&EndDate=2022-04-12'
response = urlopen(production_url)
production_data_json = json.loads(response.read())
productionapi=pd.DataFrame(production_data_json)

raw_api=productionapi.copy()
########### Convert Time ###############
raw_api['Time']= pd.to_datetime(raw_api['Time'])
raw_api['Time']=raw_api['Time'].apply(lambda t: t.strftime('%H'))
raw_api['Time']=raw_api['Time'].astype(str)
########## Convert BatchQTY ############
raw_api['BatchQty']=raw_api['BatchQty'].astype(int)
########## Convert DefectCount ############
raw_api['DefectCount']=raw_api['DefectCount'].astype(int)
########## Convert SMV ############
raw_api['SMV']=raw_api['SMV'].astype(float)
########## Convert LineNumber Into Integer #########
raw_api['LineNumber'] = raw_api['LineNumber'].str.replace('Line-','').astype(int)


raw_api=raw_api.fillna(0)
raw_api2=raw_api[['BusinessUnit', 'LineNumber','Time', 'Date','GarmentsNumber',
       'DefectName', 'DefectType', 'DefectCount']]
raw_api2

raw_api_calc=raw_api2.copy()

######################################## Total ############################################
mac=raw_api_calc.copy()
mac=mac.groupby(['Date','BusinessUnit','LineNumber']).agg({'GarmentsNumber':'nunique','DefectCount' :'sum'}).reset_index()
mac['Total Defect']=mac['DefectCount']
mac['Total Garments']=mac['GarmentsNumber']
del mac['DefectCount']
del mac['GarmentsNumber']
mac

############################################## Counting Ok Defect type ############################
mav=raw_api_calc.copy()
mav=mav[mav['DefectType']=='Ok']
mav=mav.groupby(['Date','BusinessUnit','LineNumber']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
mav['Ok']=mav['DefectType']
del mav['DefectType']
del mav['DefectCount']
mav
############################################## Counting Reject Defect type ############################
mav1=raw_api_calc.copy()
mav1=mav1[mav1['DefectType']=='Reject']
mav1=mav1.groupby(['Date','BusinessUnit','LineNumber']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
mav1['Reject']=mav1['DefectType']
del mav1['DefectType']
del mav1['DefectCount']
mav1
############################################## Counting Rework Defect type ############################
mav2=raw_api_calc.copy()
mav2=mav2[mav2['DefectType']=='Rework']
mav2=mav2.groupby(['Date','BusinessUnit','LineNumber']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
mav2['Rework']=mav2['DefectType']
del mav2['DefectType']
del mav2['DefectCount']
mav2

mao=pd.merge(mac,mav,on=['Date','BusinessUnit','LineNumber'])
mao
mao1=pd.merge(mao,mav1,on=['Date','BusinessUnit','LineNumber'],how='left')
mao1
mao2=pd.merge(mao1,mav2,on=['Date','BusinessUnit','LineNumber'],how='left')

mao2=mao2.fillna(0)

mao2['NewTotalGarments']=mao2['Total Garments']-(mao2['Reject']+mao2['Rework'])
mao2['Reject']=mao2['Reject'].astype(int)
mao2['Rework']=mao2['Rework'].astype(int)
mao2['NewTotalGarments']=mao2['NewTotalGarments'].astype(int)
del mao2['NewTotalGarments']
mao2=mao2[['Date','BusinessUnit','LineNumber','Total Defect','Total Garments','Ok', 'Reject', 'Rework']]
mao2

prod_time=raw_api.copy()

prod_time=prod_time.groupby(['Date','BusinessUnit', 'LineNumber','Time']).agg({'GarmentsNumber':'nunique'}).reset_index()
prod_time= prod_time.pivot_table('GarmentsNumber', ['Date','BusinessUnit', 'LineNumber'], 'Time').reset_index()
prod_time=prod_time.fillna(0)

prod_time['6am-7am']  =prod_time['06'].astype(int)
prod_time['7am-8am']  =prod_time['07'].astype(int)
prod_time['8am-9am']  =prod_time['08'].astype(int)
prod_time['9am-10am'] =prod_time['09'].astype(int)
prod_time['10am-11am']=prod_time['10'].astype(int)
prod_time['11am-12am']=prod_time['11'].astype(int)
prod_time['12pm-1pm'] =prod_time['12'].astype(int)
prod_time['1pm-2pm']  =prod_time['13'].astype(int)
prod_time['2pm-3pm']  =prod_time['14'].astype(int)
prod_time['3pm-4pm']  =prod_time['15'].astype(int)
prod_time['4pm-5pm']  =prod_time['16'].astype(int)
prod_time['5pm-6pm']  =prod_time['17'].astype(int)
# prod_time['6pm-7pm']  =prod_time['18'].astype(int)
prod_time['8pm-9pm']  =prod_time['19'].astype(int)
prod_time['9pm-10pm'] =prod_time['20'].astype(int)
# prod_time['11pm-12pm']=prod_time['21'].astype(int)

del prod_time['06']
del prod_time['07']
del prod_time['08']
del prod_time['09']
del prod_time['10']
del prod_time['11']
del prod_time['12']
del prod_time['13']
del prod_time['14']
del prod_time['15']
del prod_time['16']
del prod_time['17']
# del prod_time['18']
del prod_time['19']
del prod_time['20']
# del prod_time['21']

prod_merge=pd.merge(mao2,prod_time,on=['Date','BusinessUnit', 'LineNumber'])
prod_merge

############################# Forecast API & Data Conversion ########################

Forecast_excel =pd.read_excel("C:\\Users\\nadiajebin\\Desktop\\Hourly Dashboard\\Final\\Forecast_10th_April.xlsx")
forecast=Forecast_excel[['ProductionUnit','LineNumber', 'Date','PresentCadre','ForecastPcs',
                   'ForecastSAH','ForecastEff','Working_Hour','SMV']]
forecast['LineNumber'] = forecast['LineNumber'].str.replace('Line-','').astype(int)
forecast['Date']=forecast['Date'].dt.strftime('%d/%m/%Y')
forecast4apr=forecast.copy()
forecast4apr['BusinessUnit']=forecast4apr['ProductionUnit']
del forecast4apr['ProductionUnit']
forecast4apr=forecast4apr[['BusinessUnit','LineNumber','Date','PresentCadre','ForecastPcs',
                           'ForecastSAH','ForecastEff','Working_Hour','SMV']]
forecast4apr=forecast4apr.fillna(0)
forecast4apr=forecast4apr.round(2)
forecast4apr
############################# Forecast Production Merge ########################
Prod_Fore2=pd.merge(forecast4apr, prod_merge,on=['BusinessUnit','LineNumber','Date'],how='outer')
Prod_Fore2=Prod_Fore2.fillna(0)
Prod_Fore2

Prod_Fore2
Prod_Fore3=Prod_Fore2.copy()


############################# Calculations ########################

Prod_Fore2['Actual SAH']=(Prod_Fore2['SMV']*Prod_Fore2['Ok'])/60
Prod_Fore2['Actual EFF']=(Prod_Fore2['Actual SAH']*100)/(Prod_Fore2['Working_Hour']*Prod_Fore2['PresentCadre'])
Prod_Fore2=Prod_Fore2.round(2)
Prod_Fore2=Prod_Fore2.fillna(0)
Prod_Fore2

############################# Unit Wise Without Plan ########################
Prod_Fore3_sum=Prod_Fore2.groupby(['BusinessUnit','Date']).agg({'PresentCadre':'sum', 'ForecastPcs':'sum',
       'ForecastSAH':'sum','Working_Hour':'sum','Actual SAH':'sum', 'Total Defect':'sum',
       'Total Garments':'sum', 'Ok':'sum', 'Reject':'sum', 'Rework':'sum','6am-7am':'sum',
       '7am-8am':'sum', '8am-9am':'sum', '9am-10am':'sum', '10am-11am':'sum', '11am-12am':'sum', '12pm-1pm':'sum',
       '1pm-2pm':'sum', '2pm-3pm':'sum', '3pm-4pm':'sum', '4pm-5pm':'sum', '5pm-6pm':'sum', '8pm-9pm':'sum'}).reset_index()


Prod_Fore3_sum['Actual EFF']=(Prod_Fore3_sum['Actual SAH']*100)/(Prod_Fore3_sum['Working_Hour']*Prod_Fore3_sum['PresentCadre'])
#Prod_Fore3_sum['Forecast EFF']=(Prod_Fore3_sum['ForecastSAH']*100)/(Prod_Fore3_sum['Working_Hour']*Prod_Fore3_sum['PresentCadre'])
Prod_Fore3_sum['Forecast EFF']=(Prod_Fore3_sum['ForecastSAH']*100)/(24*10*26)
# Plan_Fore3_sum[]
Prod_Fore3_sum=Prod_Fore3_sum.round(2)
Prod_Fore4_sum=Prod_Fore3_sum.copy()
Prod_Fore4_sum=Prod_Fore4_sum[['BusinessUnit', 'Date', 'PresentCadre', 'Total Defect','Reject', 'Rework','Working_Hour',
               'ForecastPcs', 'ForecastSAH', 'Forecast EFF','Ok','Actual SAH','Actual EFF','6am-7am', '7am-8am', '8am-9am', '9am-10am',
               '10am-11am', '11am-12am', '12pm-1pm', '1pm-2pm', '2pm-3pm', '3pm-4pm','4pm-5pm', '5pm-6pm', '8pm-9pm']]
Prod_Fore4_sum=Prod_Fore4_sum.fillna(0)
Prod_Fore4_sum

############################# Line Wise Without Plan ########################
Prod_Fore2_line=Prod_Fore2[['BusinessUnit', 'LineNumber', 'Date', 'SMV','Working_Hour','PresentCadre','Total Defect',
                'Total Garments','Reject', 'Rework','ForecastPcs','ForecastSAH', 'ForecastEff',
                'Ok','Actual SAH', 'Actual EFF','6am-7am', '7am-8am',
       '8am-9am', '9am-10am', '10am-11am', '11am-12am', '12pm-1pm', '1pm-2pm',
       '2pm-3pm', '3pm-4pm', '4pm-5pm', '5pm-6pm', '8pm-9pm','9pm-10pm']]
Prod_Fore2_line
# '6pm-7pm', '11pm-12pm'


############################# Planned Data ########################
Planned_excel =pd.read_excel("C:\\Users\\nadiajebin\\Desktop\\Hourly Dashboard\\Final\\planned_Data_10th_1_9.xlsx")
Planned_excel
plan=Planned_excel[['BusinessUnit','LineNumber','plandate','Sum of planqty','Average of SMV']].round(2)
plan['LineNumber'] = plan['LineNumber'].str.replace('Line-','').astype(int)
plan['Date'] = plan['plandate'].str.replace('-Apr','/04/2022')
plan['Date'] = plan['Date'].str.replace('1/04/2022','01/04/2022')
plan['Date'] = plan['Date'].str.replace('2/04/2022','02/04/2022')
plan['Date'] = plan['Date'].str.replace('3/04/2022','03/04/2022')
plan['Date'] = plan['Date'].str.replace('4/04/2022','04/04/2022')
plan['Date'] = plan['Date'].str.replace('5/04/2022','05/04/2022')
plan['Date'] = plan['Date'].str.replace('6/04/2022','06/04/2022')
plan['Date'] = plan['Date'].str.replace('7/04/2022','07/04/2022')
plan['Date'] = plan['Date'].str.replace('8/04/2022','08/04/2022')
plan['Date'] = plan['Date'].str.replace('9/04/2022','09/04/2022')

del plan['plandate']
plan=plan[['Date','BusinessUnit','LineNumber', 'Sum of planqty', 'Average of SMV']]

plan['Planned SAH']=(plan['Sum of planqty']*plan['Average of SMV'])/60
plan



#################################### Merge Plan and Forecast-Actual #######################################
''' Line Wise'''

Prod_Fore_plan_line=pd.merge(plan, Prod_Fore2_line,on=['BusinessUnit','LineNumber','Date'],how='outer')
Prod_Fore_plan_line['Planned EFF']=(Prod_Fore_plan_line['Planned SAH']*100)/(10*24)
Prod_Fore_plan_line=Prod_Fore_plan_line[['Date', 'BusinessUnit', 'LineNumber','Total Defect', 'Total Garments', 'Reject', 'Rework',
                    'SMV', 'Working_Hour', 'PresentCadre','Sum of planqty', 'Planned SAH','Planned EFF',
                    'ForecastPcs','ForecastSAH', 'ForecastEff', 'Ok', 'Actual SAH', 'Actual EFF',
                    '6am-7am', '7am-8am', '8am-9am', '9am-10am', '10am-11am', '11am-12am',
                    '12pm-1pm', '1pm-2pm', '2pm-3pm', '3pm-4pm', '4pm-5pm', '5pm-6pm','8pm-9pm', '9pm-10pm']]
# '6pm-7pm', '11pm-12pm'
Prod_Fore_plan_line=Prod_Fore_plan_line.round(2)
Prod_Fore_plan_line=Prod_Fore_plan_line.fillna(0)
Prod_Fore_plan_line
excel1 = Prod_Fore_plan_line.to_excel('C:\\Users\\nadiajebin\\Desktop\\send\\Line Wise Report.xlsx')


''' Unit Wise'''

plan_unit=plan.copy()

plan_unit=plan_unit.groupby(['Date','BusinessUnit']).agg({'Sum of planqty':'sum', 'Planned SAH':'sum'}).reset_index()
plan_unit['Planned EFF']=(plan_unit['Planned SAH']*100)/(24*10*26)

Prod_Fore_plan_Unit=pd.merge(plan_unit, Prod_Fore3_sum,on=['BusinessUnit','Date'],how='outer')
Prod_Fore_plan_Unit=Prod_Fore_plan_Unit.round(2)
Prod_Fore_plan_Unit=Prod_Fore_plan_Unit[['Date', 'BusinessUnit','Sum of planqty','Working_Hour','Planned SAH', 'Planned EFF',
       'ForecastPcs', 'ForecastSAH', 'Forecast EFF', 'Ok','Actual SAH','Actual EFF','6am-7am', '7am-8am', '8am-9am',
       '9am-10am', '10am-11am','11am-12am', '12pm-1pm', '1pm-2pm', '2pm-3pm', '3pm-4pm', '4pm-5pm',
       '5pm-6pm', '8pm-9pm']]
Prod_Fore_plan_Unit=Prod_Fore_plan_Unit.fillna(0)

excel2 = Prod_Fore_plan_Unit.to_excel('C:\\Users\\nadiajebin\\Desktop\\send\\Unit Wise Plan Report.xlsx')

############################ Report With STYLE, PO,COLOR ######################################
'''  Access Production API and Convert Data '''
raw_api=productionapi.copy()
########### Convert Time ###############
raw_api['Time']= pd.to_datetime(raw_api['Time'])
raw_api['Time']=raw_api['Time'].apply(lambda t: t.strftime('%H'))
raw_api['Time']=raw_api['Time'].astype(str)
########## Convert BatchQTY ############
raw_api['BatchQty']=raw_api['BatchQty'].astype(int)
########## Convert DefectCount ############
raw_api['DefectCount']=raw_api['DefectCount'].astype(int)
########## Convert SMV ############
raw_api['SMV']=raw_api['SMV'].astype(float)
########## Convert LineNumber Into Integer #########
raw_api['LineNumber'] = raw_api['LineNumber'].str.replace('Line-','').astype(int)

raw_api

raw_api=raw_api.fillna(0)
raw_api2=raw_api[['BusinessUnit', 'LineNumber','Time', 'Date','StyleSubCat', 'PoNumber','Color','GarmentsNumber',
       'Size','DefectName', 'DefectType', 'DefectCount']]
raw_api2

raw_api_calc=raw_api2.copy()

######################################## Total ############################################
mac=raw_api_calc.copy()
mac=mac.groupby(['Date','BusinessUnit','LineNumber','StyleSubCat', 'PoNumber','Color','Size']).agg({'GarmentsNumber':'nunique','DefectCount' :'sum'}).reset_index()
mac['Total Defect']=mac['DefectCount']
mac['Total Garments']=mac['GarmentsNumber']
del mac['DefectCount']
del mac['GarmentsNumber']
mac

############################################## Counting Ok Defect type ############################
mav=raw_api_calc.copy()
mav=mav[mav['DefectType']=='Ok']
mav=mav.groupby(['Date','BusinessUnit','LineNumber','StyleSubCat', 'PoNumber','Color','Size']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
mav['Ok']=mav['DefectType']
del mav['DefectType']
del mav['DefectCount']
mav
############################################## Counting Reject Defect type ############################
mav1=raw_api_calc.copy()
mav1=mav1[mav1['DefectType']=='Reject']
mav1=mav1.groupby(['Date','BusinessUnit','LineNumber','StyleSubCat', 'PoNumber','Color','Size']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
mav1['Reject']=mav1['DefectType']
del mav1['DefectType']
del mav1['DefectCount']
mav1
############################################## Counting Rework Defect type ############################
mav2=raw_api_calc.copy()
mav2=mav2[mav2['DefectType']=='Rework']
mav2=mav2.groupby(['Date','BusinessUnit','LineNumber','StyleSubCat', 'PoNumber','Color','Size']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
mav2['Rework']=mav2['DefectType']
del mav2['DefectType']
del mav2['DefectCount']
mav2

############################################## Merging All Three Tables ############################
mao=pd.merge(mac,mav,on=['Date','BusinessUnit','LineNumber','StyleSubCat', 'PoNumber','Color','Size'],how='outer')
mao
mao1=pd.merge(mao,mav1,on=['Date','BusinessUnit','LineNumber','StyleSubCat', 'PoNumber','Color','Size'],how='outer')
mao1
mao2=pd.merge(mao1,mav2,on=['Date','BusinessUnit','LineNumber','StyleSubCat', 'PoNumber','Color','Size'],how='outer')

mao2=mao2.fillna(0)


############################################## Calculation and Data Type Conversion ############################
mao2['NewTotalGarments']=mao2['Total Garments']-(mao2['Reject']+mao2['Rework'])
mao2['Reject']=mao2['Reject'].astype(int)
mao2['Rework']=mao2['Rework'].astype(int)
mao2['Ok']=mao2['Ok'].astype(int)
mao2['NewTotalGarments']=mao2['NewTotalGarments'].astype(int)
del mao2['NewTotalGarments']
mao2=mao2[['Date','BusinessUnit','LineNumber','StyleSubCat', 'PoNumber','Color','Size','Total Defect','Total Garments','Ok', 'Reject', 'Rework']]
mao2

############################################## Pivoting Time  #####################################################
prod_time=raw_api.copy()
prod_time=prod_time.groupby(['Date','BusinessUnit', 'LineNumber','StyleSubCat', 'PoNumber',
                             'Color','Size','Time']).agg({'GarmentsNumber':'nunique'}).reset_index()
prod_time= prod_time.pivot_table('GarmentsNumber', ['Date','BusinessUnit', 'LineNumber','StyleSubCat',
                                                    'PoNumber','Color','Size'], 'Time').reset_index()
prod_time=prod_time.fillna(0)
prod_time

############################################## Rename Hour Column  #####################################################
prod_time['6am-7am']  =prod_time['06'].astype(int)
prod_time['7am-8am']  =prod_time['07'].astype(int)
prod_time['8am-9am']  =prod_time['08'].astype(int)
prod_time['9am-10am'] =prod_time['09'].astype(int)
prod_time['10am-11am']=prod_time['10'].astype(int)
prod_time['11am-12am']=prod_time['11'].astype(int)
prod_time['12pm-1pm'] =prod_time['12'].astype(int)
prod_time['1pm-2pm']  =prod_time['13'].astype(int)
prod_time['2pm-3pm']  =prod_time['14'].astype(int)
prod_time['3pm-4pm']  =prod_time['15'].astype(int)
prod_time['4pm-5pm']  =prod_time['16'].astype(int)
prod_time['5pm-6pm']  =prod_time['17'].astype(int)
# prod_time['6pm-7pm']  =prod_time['18'].astype(int)
prod_time['8pm-9pm']  =prod_time['19'].astype(int)
prod_time['9pm-10pm'] =prod_time['20'].astype(int)
# prod_time['11pm-12pm']=prod_time['21'].astype(int)

del prod_time['06']
del prod_time['07']
del prod_time['08']
del prod_time['09']
del prod_time['10']
del prod_time['11']
del prod_time['12']
del prod_time['13']
del prod_time['14']
del prod_time['15']
del prod_time['16']
del prod_time['17']
# del prod_time['18']
del prod_time['19']
del prod_time['20']
# del prod_time['21']

prod_merge=pd.merge(mao2,prod_time,on=['Date','BusinessUnit', 'LineNumber','StyleSubCat',
                                                    'PoNumber','Color','Size'], how='outer')
prod_merge

############################ Forecast Data ################################
Forecast_excel =pd.read_excel("C:\\Users\\nadiajebin\\Desktop\\Hourly Dashboard\\Final\\Forecast_10th_April.xlsx")
forecast=Forecast_excel[['ProductionUnit','LineNumber', 'Date','PresentCadre','ForecastPcs',
                   'ForecastSAH','ForecastEff','Working_Hour','SMV']]
forecast['LineNumber'] = forecast['LineNumber'].str.replace('Line-','').astype(int)
forecast['Date']=forecast['Date'].dt.strftime('%d/%m/%Y')
forecast4apr=forecast.copy()
forecast4apr['BusinessUnit']=forecast4apr['ProductionUnit']
del forecast4apr['ProductionUnit']
forecast4apr=forecast4apr[['BusinessUnit','LineNumber','Date','PresentCadre','ForecastPcs',
                           'ForecastSAH','ForecastEff','Working_Hour','SMV']]
forecast4apr=forecast4apr.fillna(0)
forecast4apr=forecast4apr.round(2)
forecast4apr

Prod_Fore2=pd.merge(forecast4apr, prod_merge,on=['BusinessUnit','LineNumber','Date'],how='outer')
Prod_Fore2=Prod_Fore2.fillna(0)
Prod_Fore2

Prod_Fore2
Prod_Fore3=Prod_Fore2.copy()
Prod_Fore2['Actual SAH']=(Prod_Fore2['SMV']*Prod_Fore2['Ok'])/60
Prod_Fore2['Actual EFF']=(Prod_Fore2['Actual SAH']*100)/(Prod_Fore2['Working_Hour']*Prod_Fore2['PresentCadre'])
Prod_Fore2=Prod_Fore2.round(2)
Prod_Fore2=Prod_Fore2.fillna(0)
Prod_Fore2.columns

Prod_Fore3_sum=Prod_Fore2.groupby(['BusinessUnit','Date','StyleSubCat',
       'PoNumber', 'Color', 'Size']).agg({'PresentCadre':'sum', 'ForecastPcs':'sum',
       'ForecastSAH':'sum','Working_Hour':'sum','Actual SAH':'sum', 'Total Defect':'sum',
       'Total Garments':'sum', 'Ok':'sum', 'Reject':'sum', 'Rework':'sum','6am-7am':'sum',
       '7am-8am':'sum', '8am-9am':'sum', '9am-10am':'sum', '10am-11am':'sum', '11am-12am':'sum', '12pm-1pm':'sum',
       '1pm-2pm':'sum', '2pm-3pm':'sum', '3pm-4pm':'sum', '4pm-5pm':'sum', '5pm-6pm':'sum', '8pm-9pm':'sum'}).reset_index()


Prod_Fore3_sum['Actual EFF']=(Prod_Fore3_sum['Actual SAH']*100)/(Prod_Fore3_sum['Working_Hour']*Prod_Fore3_sum['PresentCadre'])
Prod_Fore3_sum['Forecast EFF']=(Prod_Fore3_sum['ForecastSAH']*100)/(Prod_Fore3_sum['Working_Hour']*Prod_Fore3_sum['PresentCadre'])

Prod_Fore3_sum=Prod_Fore3_sum.round(2)
Prod_Fore3_sum=Prod_Fore3_sum.fillna(0)
Prod_Fore4_sum=Prod_Fore3_sum.copy()
Prod_Fore4_sum=Prod_Fore4_sum[['BusinessUnit', 'Date', 'StyleSubCat', 'PoNumber', 'Color', 'Size','Total Defect','Reject',
       'Rework', 'Working_Hour', 'PresentCadre', 'ForecastPcs', 'ForecastSAH','Forecast EFF','Ok','Actual SAH','Actual EFF',
       '6am-7am', '7am-8am', '8am-9am', '9am-10am', '10am-11am','11am-12am', '12pm-1pm', '1pm-2pm', '2pm-3pm', '3pm-4pm',
       '4pm-5pm','5pm-6pm', '8pm-9pm']]
Prod_Fore4_sum=Prod_Fore4_sum.fillna(0)
Prod_Fore4_sum


################################### Plan Data #########################

Planned_excel =pd.read_excel("C:\\Users\\nadiajebin\\Desktop\\Hourly Dashboard\\Final\\planned_Data_10th_1_9.xlsx")
Planned_excel
plan=Planned_excel[['BusinessUnit','LineNumber','plandate','Sum of planqty','Average of SMV']].round(2)
########## Convert LineNumber Into Integer #########
plan['LineNumber'] = plan['LineNumber'].str.replace('Line-','').astype(int)

plan['Date'] = plan['plandate'].str.replace('-Apr','/04/2022')
plan['Date'] = plan['Date'].str.replace('1/04/2022','01/04/2022')
plan['Date'] = plan['Date'].str.replace('2/04/2022','02/04/2022')
plan['Date'] = plan['Date'].str.replace('3/04/2022','03/04/2022')
plan['Date'] = plan['Date'].str.replace('4/04/2022','04/04/2022')
plan['Date'] = plan['Date'].str.replace('5/04/2022','05/04/2022')
plan['Date'] = plan['Date'].str.replace('6/04/2022','06/04/2022')
plan['Date'] = plan['Date'].str.replace('7/04/2022','07/04/2022')
plan['Date'] = plan['Date'].str.replace('8/04/2022','08/04/2022')
plan['Date'] = plan['Date'].str.replace('9/04/2022','09/04/2022')

del plan['plandate']
plan=plan[['Date','BusinessUnit','LineNumber', 'Sum of planqty', 'Average of SMV']]

plan['Planned SAH']=(plan['Sum of planqty']*plan['Average of SMV'])/60
plan

#################################### Unit Wise Report ###########################
plan_unit=plan.copy()

plan_unit=plan_unit.groupby(['Date','BusinessUnit']).agg({'Sum of planqty':'sum', 'Planned SAH':'sum'}).reset_index()
plan_unit['Planned EFF']=(plan_unit['Planned SAH']*100)/(24*10*26)

Prod_Fore_plan_Unit=pd.merge(plan_unit, Prod_Fore3_sum,on=['BusinessUnit','Date'],how='outer')
Prod_Fore_plan_Unit=Prod_Fore_plan_Unit.round(2)
Prod_Fore_plan_Unit=Prod_Fore_plan_Unit.fillna(0)

Prod_Fore_plan_Unit['Planned Pcs']=Prod_Fore_plan_Unit['Sum of planqty'].astype(int)
del Prod_Fore_plan_Unit['Sum of planqty']
Prod_Fore_plan_Unit['ForecastPcs']=Prod_Fore_plan_Unit['ForecastPcs'].astype(int)
Prod_Fore_plan_Unit['Ok']=Prod_Fore_plan_Unit['Ok'].astype(int)
Prod_Fore_plan_Unit['Working_Hour']  =Prod_Fore_plan_Unit['Working_Hour'].astype(int)

Prod_Fore_plan_Unit=Prod_Fore_plan_Unit[['Date', 'BusinessUnit','StyleSubCat', 'PoNumber', 'Color', 'Size','Planned Pcs','Working_Hour','Planned SAH', 'Planned EFF',
       'ForecastPcs', 'ForecastSAH', 'Forecast EFF', 'Ok','Actual SAH','Actual EFF','6am-7am', '7am-8am', '8am-9am',
       '9am-10am', '10am-11am','11am-12am', '12pm-1pm', '1pm-2pm', '2pm-3pm', '3pm-4pm', '4pm-5pm',
       '5pm-6pm', '8pm-9pm']]
Prod_Fore_plan_Unit.iloc[:,16:]  =Prod_Fore_plan_Unit.iloc[:,16:].astype(int)

Prod_Fore_plan_Unit.iloc[:,12:]



excel3 = Prod_Fore_plan_Unit.to_excel('C:\\Users\\nadiajebin\\Desktop\\send\\Unit Wise Plan Report with Style.xlsx')

############################## Line Wise Report ################################
Prod_Fore2_line=Prod_Fore2[['BusinessUnit', 'LineNumber', 'Date','StyleSubCat','PoNumber', 'Color','Size', 'SMV','Working_Hour',
                'PresentCadre','Total Defect','Reject', 'Rework','ForecastPcs','ForecastSAH', 'ForecastEff',
                'Ok','Actual SAH', 'Actual EFF','6am-7am', '7am-8am','8am-9am', '9am-10am', '10am-11am',
                '11am-12am', '12pm-1pm', '1pm-2pm','2pm-3pm', '3pm-4pm', '4pm-5pm', '5pm-6pm', '8pm-9pm','9pm-10pm']]
# '6pm-7pm','11pm-12pm'
Prod_Fore2_line
Prod_Fore_plan_line=pd.merge(plan, Prod_Fore2_line,on=['BusinessUnit','LineNumber','Date'],how='outer')
Prod_Fore_plan_line
Prod_Fore_plan_line['Planned EFF']=(Prod_Fore_plan_line['Planned SAH']*100)/(10*24)
Prod_Fore_plan_line=Prod_Fore_plan_line.round(2)
Prod_Fore_plan_line=Prod_Fore_plan_line.fillna(0)
Prod_Fore_plan_line['Planned Pcs']=Prod_Fore_plan_line['Sum of planqty'].astype(int)
del Prod_Fore_plan_line['Sum of planqty']
Prod_Fore_plan_line['Total Defect']=Prod_Fore_plan_line['Total Defect'].astype(int)
Prod_Fore_plan_line['Reject']=Prod_Fore_plan_line['Reject'].astype(int)
Prod_Fore_plan_line['Rework']=Prod_Fore_plan_line['Rework'].astype(int)
Prod_Fore_plan_line['Working_Hour']=Prod_Fore_plan_line['Working_Hour'].astype(int)
Prod_Fore_plan_line['ForecastPcs']=Prod_Fore_plan_line['ForecastPcs'].astype(int)
Prod_Fore_plan_line['Ok']=Prod_Fore_plan_line['Ok'].astype(int)

Prod_Fore_plan_line=Prod_Fore_plan_line[['Date', 'BusinessUnit', 'LineNumber','StyleSubCat','PoNumber', 'Color','Size',
                    'Total Defect','Reject', 'Rework','SMV', 'Working_Hour', 'PresentCadre',
                    'Planned Pcs', 'Planned SAH','Planned EFF',
                    'ForecastPcs','ForecastSAH', 'ForecastEff',
                    'Ok', 'Actual SAH', 'Actual EFF',
                    '6am-7am', '7am-8am', '8am-9am', '9am-10am', '10am-11am', '11am-12am','12pm-1pm', '1pm-2pm', '2pm-3pm',
                    '3pm-4pm', '4pm-5pm', '5pm-6pm','8pm-9pm', '9pm-10pm']]
# '6pm-7pm', '11pm-12pm'
Prod_Fore_plan_line.iloc[:,22:]=Prod_Fore_plan_line.iloc[:,22:].astype(int)

Prod_Fore_plan_line.iloc[:,17:]

excel4 = Prod_Fore_plan_line.to_excel('C:\\Users\\nadiajebin\\Desktop\\send\\Line Wise Report with Style.xlsx')

############################## App #############################
Prod_Fore2_line=Prod_Fore2[['BusinessUnit', 'LineNumber', 'Date','StyleSubCat','PoNumber', 'Color','Size', 'SMV','Working_Hour',
                'PresentCadre','Total Defect','Reject', 'Rework','ForecastPcs','ForecastSAH', 'ForecastEff',
                'Ok','Actual SAH', 'Actual EFF','6am-7am', '7am-8am','8am-9am', '9am-10am', '10am-11am',
                '11am-12am', '12pm-1pm', '1pm-2pm','2pm-3pm', '3pm-4pm', '4pm-5pm', '5pm-6pm', '8pm-9pm','9pm-10pm']]
# '6pm-7pm','11pm-12pm'
Prod_Fore2_line
Prod_Fore_plan_line=pd.merge(plan, Prod_Fore2_line,on=['BusinessUnit','LineNumber','Date'],how='outer')
Prod_Fore_plan_line
Prod_Fore_plan_line['Planned EFF']=(Prod_Fore_plan_line['Planned SAH']*100)/(10*24)
Prod_Fore_plan_line=Prod_Fore_plan_line.round(2)
Prod_Fore_plan_line=Prod_Fore_plan_line.fillna(0)
Prod_Fore_plan_line['Planned Pcs']=Prod_Fore_plan_line['Sum of planqty'].astype(int)
del Prod_Fore_plan_line['Sum of planqty']
Prod_Fore_plan_line['Total Defect']=Prod_Fore_plan_line['Total Defect'].astype(int)
Prod_Fore_plan_line['Reject']=Prod_Fore_plan_line['Reject'].astype(int)
Prod_Fore_plan_line['Rework']=Prod_Fore_plan_line['Rework'].astype(int)
Prod_Fore_plan_line['Working_Hour']=Prod_Fore_plan_line['Working_Hour'].astype(int)
Prod_Fore_plan_line['ForecastPcs']=Prod_Fore_plan_line['ForecastPcs'].astype(int)
Prod_Fore_plan_line['Ok']=Prod_Fore_plan_line['Ok'].astype(int)

Prod_Fore_plan_line=Prod_Fore_plan_line[['Date', 'BusinessUnit', 'LineNumber','StyleSubCat','PoNumber', 'Color','Size',
                    'Total Defect','Reject', 'Rework','SMV', 'Working_Hour', 'PresentCadre',
                    'Planned Pcs', 'Planned SAH','Planned EFF',
                    'ForecastPcs','ForecastSAH', 'ForecastEff',
                    'Ok', 'Actual SAH', 'Actual EFF',
                    '6am-7am', '7am-8am', '8am-9am', '9am-10am', '10am-11am', '11am-12am','12pm-1pm', '1pm-2pm', '2pm-3pm',
                    '3pm-4pm', '4pm-5pm', '5pm-6pm','8pm-9pm', '9pm-10pm']]
# '6pm-7pm', '11pm-12pm'
Prod_Fore_plan_line.iloc[:,22:]=Prod_Fore_plan_line.iloc[:,22:].astype(int)

Prod_Fore_plan_line.iloc[:,17:]

excel4 = Prod_Fore_plan_line.to_excel('C:\\Users\\nadiajebin\\Desktop\\send\\Line Wise Report with Style.xlsx')

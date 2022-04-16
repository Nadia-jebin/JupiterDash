import pandas as pd

from dash import Dash, dcc, html
from dash.dependencies import Input,State, Output

import dash_bootstrap_components as dbc

from dash import callback_context
from flask import Flask, send_from_directory
import json
from urllib.request import urlopen

production_url='http://10.12.13.164:8110/api/BasicQMSData/BirichinaDataForReport?StartDate=2022-04-02&EndDate=2022-04-13'
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

raw_api['Unique Garments']=raw_api['GarmentsNumber']
raw_api1=raw_api[['BusinessUnit', 'LineNumber','Date','StyleSubCat', 'PoNumber','Color','Size','Unique Garments','GarmentsNumber',
       'DefectName','DefectType', 'DefectCount','ModuleName']]

raw_api1

raw_api_sewing=raw_api1.copy()
raw_api_sewing=raw_api1[raw_api1['ModuleName']=='Sewing']
raw_api_sewing

raw_api_calc=raw_api1.copy()

######################################## Total ############################################
mac=raw_api_calc.copy()
mac=mac.groupby(['Date','BusinessUnit','LineNumber','StyleSubCat','PoNumber',
                           'Color']).agg({'GarmentsNumber':'nunique','DefectCount' :'sum'}).reset_index()
mac['Total Defect']=mac['DefectCount']
mac['Total Garments']=mac['GarmentsNumber']
del mac['DefectCount']
del mac['GarmentsNumber']
mac

############################################## Counting Ok Defect type ############################
mav=raw_api_calc.copy()
mav=mav[mav['DefectType']=='Ok']
mav=mav.groupby(['Date','BusinessUnit','LineNumber','StyleSubCat','PoNumber',
                           'Color']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
mav['Ok']=mav['DefectType']
del mav['DefectType']
del mav['DefectCount']
mav
############################################## Counting Reject Defect type ############################
mav1=raw_api_calc.copy()
mav1=mav1[mav1['DefectType']=='Reject']
mav1=mav1.groupby(['Date','BusinessUnit','LineNumber','StyleSubCat','PoNumber',
                           'Color']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
mav1['Reject']=mav1['DefectType']
del mav1['DefectType']
del mav1['DefectCount']
mav1
############################################## Counting Rework Defect type ############################
mav2=raw_api_calc.copy()
mav2=mav2[mav2['DefectType']=='Rework']
mav2=mav2.groupby(['Date','BusinessUnit','LineNumber','StyleSubCat','PoNumber',
                           'Color']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
mav2['Rework']=mav2['DefectType']
del mav2['DefectType']
del mav2['DefectCount']
mav2

mao=pd.merge(mac,mav,on=['Date','BusinessUnit','LineNumber','StyleSubCat','PoNumber',
                           'Color'])
mao
mao1=pd.merge(mao,mav1,on=['Date','BusinessUnit','LineNumber','StyleSubCat','PoNumber',
                           'Color'],how='left')
mao1
mao2=pd.merge(mao1,mav2,on=['Date','BusinessUnit','LineNumber','StyleSubCat','PoNumber',
                           'Color'],how='left')

mao2=mao2.fillna(0)

mao2['NewTotalGarments']=mao2['Total Garments']-(mao2['Reject']+mao2['Rework'])
mao2['Reject']=mao2['Reject'].astype(int)
mao2['Rework']=mao2['Rework'].astype(int)
mao2['NewTotalGarments']=mao2['NewTotalGarments'].astype(int)
del mao2['NewTotalGarments']
mao2=mao2[['Date','BusinessUnit','LineNumber','StyleSubCat','PoNumber','Color','Total Defect',
           'Total Garments','Ok', 'Reject', 'Rework']]
mao2['DHU%']=(mao2['Total Defect']*100)/mao2['Ok']
mao2['Reject%']=(mao2['Reject']*100)/mao2['Ok']
mao2=mao2.round(2)
mao2

################### Only Module == Sewing ######################
raw_api_calc=raw_api_sewing.copy()

######################################## Total ############################################
mac=raw_api_calc.copy()
mac=mac.groupby(['Date','BusinessUnit','LineNumber','StyleSubCat','PoNumber',
                           'Color']).agg({'GarmentsNumber':'nunique','DefectCount' :'sum'}).reset_index()
mac['Total Defect']=mac['DefectCount']
mac['Total Garments']=mac['GarmentsNumber']
del mac['DefectCount']
del mac['GarmentsNumber']
mac

############################################## Counting Ok Defect type ############################
mav=raw_api_calc.copy()
mav=mav[mav['DefectType']=='Ok']
mav=mav.groupby(['Date','BusinessUnit','LineNumber','StyleSubCat','PoNumber',
                           'Color']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
mav['Ok']=mav['DefectType']
del mav['DefectType']
del mav['DefectCount']
mav
############################################## Counting Reject Defect type ############################
mav1=raw_api_calc.copy()
mav1=mav1[mav1['DefectType']=='Reject']
mav1=mav1.groupby(['Date','BusinessUnit','LineNumber','StyleSubCat','PoNumber',
                           'Color']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
mav1['Reject']=mav1['DefectType']
del mav1['DefectType']
del mav1['DefectCount']
mav1
############################################## Counting Rework Defect type ############################
mav2=raw_api_calc.copy()
mav2=mav2[mav2['DefectType']=='Rework']
mav2=mav2.groupby(['Date','BusinessUnit','LineNumber','StyleSubCat','PoNumber',
                           'Color']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
mav2['Rework']=mav2['DefectType']
del mav2['DefectType']
del mav2['DefectCount']
mav2

mao=pd.merge(mac,mav,on=['Date','BusinessUnit','LineNumber','StyleSubCat','PoNumber',
                           'Color'])
mao
mao1=pd.merge(mao,mav1,on=['Date','BusinessUnit','LineNumber','StyleSubCat','PoNumber',
                           'Color'],how='left')
mao1
mao2=pd.merge(mao1,mav2,on=['Date','BusinessUnit','LineNumber','StyleSubCat','PoNumber',
                           'Color'],how='left')

mao2=mao2.fillna(0)

mao2['NewTotalGarments']=mao2['Total Garments']-(mao2['Reject']+mao2['Rework'])
mao2['Reject']=mao2['Reject'].astype(int)
mao2['Rework']=mao2['Rework'].astype(int)
mao2['NewTotalGarments']=mao2['NewTotalGarments'].astype(int)
del mao2['NewTotalGarments']
mao2=mao2[['Date','BusinessUnit','LineNumber','StyleSubCat','PoNumber','Color','Total Defect',
           'Total Garments','Ok', 'Reject', 'Rework']]
mao2['DHU%']=(mao2['Total Defect']*100)/mao2['Ok']
mao2['Reject%']=(mao2['Reject']*100)/mao2['Ok']
mao2=mao2.round(2)
mao2

mao2['Check Qty']=mao2['Ok']
mao2['Defect Qty']=mao2['Total Defect']
del mao2['Ok']
del mao2['Total Defect']

mao2=mao2[['Date','BusinessUnit','LineNumber','StyleSubCat','PoNumber','Color','Total Garments','Rework',
           'Check Qty','Defect Qty','DHU%','Check Qty','Reject','Reject%']]
mao2

#################################### Report 14 ##################################
'''################# Table 1 ###############'''
report_14=raw_api1.copy()

######################################## Total ############################################
allD=report_14.copy()
allD=allD.groupby(['Date','BusinessUnit']).agg({'GarmentsNumber':'nunique','DefectCount' :'sum'}).reset_index()
allD['Total Defect']=allD['DefectCount']
allD['Total Garments']=allD['GarmentsNumber']
del allD['DefectCount']
del allD['GarmentsNumber']
allD

############################################## Counting Ok Defect type ############################
Ok_count_14=report_14.copy()
Ok_count_14=Ok_count_14[Ok_count_14['DefectType']=='Ok']
Ok_count_14=Ok_count_14.groupby(['Date','BusinessUnit']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
Ok_count_14['Ok']=Ok_count_14['DefectType']
del Ok_count_14['DefectType']
del Ok_count_14['DefectCount']
Ok_count_14
############################################## Counting Reject Defect type ############################
reject_count_14=report_14.copy()
reject_count_14=reject_count_14[reject_count_14['DefectType']=='Reject']
reject_count_14=reject_count_14.groupby(['Date','BusinessUnit']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
reject_count_14['Reject']=reject_count_14['DefectType']
del reject_count_14['DefectType']
del reject_count_14['DefectCount']
reject_count_14
############################################## Counting Rework Defect type ############################
rework_14=report_14.copy()
rework_14=rework_14[rework_14['DefectType']=='Rework']
rework_14=rework_14.groupby(['Date','BusinessUnit']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
rework_14['Rework']=rework_14['DefectType']
del rework_14['DefectType']
del rework_14['DefectCount']
rework_14

allD_ok=pd.merge(allD,Ok_count_14,on=['Date','BusinessUnit'])
allD_ok
allD_ok_rej=pd.merge(allD_ok,reject_count_14,on=['Date','BusinessUnit'],how='left')
allD_ok_rej
allD_ok_rej_rw=pd.merge(allD_ok_rej,rework_14,on=['Date','BusinessUnit'],how='left')

allD_ok_rej_rw=allD_ok_rej_rw.fillna(0)

allD_ok_rej_rw['NewTotalGarments']=allD_ok_rej_rw['Total Garments']-(allD_ok_rej_rw['Reject']+allD_ok_rej_rw['Rework'])
allD_ok_rej_rw['Reject']=allD_ok_rej_rw['Reject'].astype(int)
allD_ok_rej_rw['Rework']=allD_ok_rej_rw['Rework'].astype(int)
allD_ok_rej_rw['NewTotalGarments']=allD_ok_rej_rw['NewTotalGarments'].astype(int)
del allD_ok_rej_rw['NewTotalGarments']
allD_ok_rej_rw=allD_ok_rej_rw[['Date','BusinessUnit','Total Defect','Total Garments','Ok', 'Reject', 'Rework']]
allD_ok_rej_rw

allD_ok_rej_rw['Check Qty']=allD_ok_rej_rw['Ok']
allD_ok_rej_rw['Defect Qty']=allD_ok_rej_rw['Total Defect']
del allD_ok_rej_rw['Ok']
del allD_ok_rej_rw['Total Defect']
allD_ok_rej_rw['DHU%']=(allD_ok_rej_rw['Defect Qty']*100)/allD_ok_rej_rw['Check Qty']
allD_ok_rej_rw['Reject%']=(allD_ok_rej_rw['Reject']*100)/allD_ok_rej_rw['Check Qty']
allD_ok_rej_rw=allD_ok_rej_rw.round(2)
allD_ok_rej_rw

Forecast_excel =pd.read_excel("C:\\Users\\nadiajebin\\Desktop\\Hourly Dashboard\\Final\\Forecast_10th_April.xlsx")
# forecast=Forecast_excel[['ProductionUnit','LineNumber', 'Date','PresentCadre','ForecastPcs',
#                    'ForecastSAH','ForecastEff','Working_Hour','SMV']]
forecast=Forecast_excel.groupby(['Date','ProductionUnit']).agg({'ForecastPcs' :'sum'}).reset_index()
# forecast['LineNumber'] = forecast['LineNumber'].str.replace('Line-','').astype(int)
forecast['Date']=forecast['Date'].dt.strftime('%d/%m/%Y')
forecast4apr=forecast.copy()
forecast4apr['BusinessUnit']=forecast4apr['ProductionUnit']
del forecast4apr['ProductionUnit']
forecast4apr=forecast4apr[['Date','BusinessUnit','ForecastPcs']]
forecast4apr=forecast4apr.fillna(0)
forecast4apr=forecast4apr.round(2)
forecast4apr

Prod_Fore2=pd.merge(forecast4apr, allD_ok_rej_rw,on=['Date','BusinessUnit'],how='outer')
Prod_Fore2=Prod_Fore2.fillna(0)
Prod_Fore2

Prod_Fore2


Prod_Fore2=Prod_Fore2.round(2)
Prod_Fore2=Prod_Fore2.fillna(0)

Prod_Fore2

Planned_excel =pd.read_excel("C:\\Users\\nadiajebin\\Desktop\\Hourly Dashboard\\Final\\planned_Data_10th_1_9.xlsx")
Planned_excel
Planned_excel['Planned Pcs']=Planned_excel['Sum of planqty']
plan=Planned_excel[['BusinessUnit','plandate','Planned Pcs','Average of SMV']].round(2)
# plan['LineNumber'] = plan['LineNumber'].str.replace('Line-','').astype(int)
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
plan=plan[['Date','BusinessUnit','Planned Pcs']]
plan

plan_unit=plan.copy()

plan_unit=plan_unit.groupby(['Date','BusinessUnit']).agg({'Planned Pcs':'sum'}).reset_index()
Prod_Fore_plan_Unit=pd.merge(plan_unit, Prod_Fore2,on=['Date','BusinessUnit'],how='outer')
Prod_Fore_plan_Unit=Prod_Fore_plan_Unit.round(2)
Prod_Fore_plan_Unit=Prod_Fore_plan_Unit.fillna(0)
Prod_Fore_plan_Unit.loc[:,"Planned Pcs":"Defect Qty"]=Prod_Fore_plan_Unit.loc[:,"Planned Pcs":"Defect Qty"].astype(int)
Prod_Fore_plan_Unit=Prod_Fore_plan_Unit

###################################### Table 2 ###########################

report_14=raw_api1.copy()

######################################## Total ############################################
TallD=report_14.copy()
TallD=TallD.groupby(['Date','BusinessUnit','DefectName']).agg({'GarmentsNumber':'nunique','DefectCount' :'sum'}).reset_index()
TallD['Total Defect']=TallD['DefectCount']
TallD['Total Garments']=TallD['GarmentsNumber']
del TallD['DefectCount']
del TallD['GarmentsNumber']
TallD

############################################## Counting Ok Defect type ############################
TOk_count_14=report_14.copy()
TOk_count_14=TOk_count_14[TOk_count_14['DefectType']=='Ok']
TOk_count_14=TOk_count_14.groupby(['Date','BusinessUnit','DefectName']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
TOk_count_14['Ok']=TOk_count_14['DefectType']
del TOk_count_14['DefectType']
del TOk_count_14['DefectCount']
TOk_count_14
############################################## Counting Reject Defect type ############################
Treject_count_14=report_14.copy()
Treject_count_14=Treject_count_14[Treject_count_14['DefectType']=='Reject']
Treject_count_14=Treject_count_14.groupby(['Date','BusinessUnit','DefectName']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
Treject_count_14['Reject']=Treject_count_14['DefectType']
del Treject_count_14['DefectType']
del Treject_count_14['DefectCount']
Treject_count_14
############################################## Counting Rework Defect type ############################
Trework_14=report_14.copy()
Trework_14=Trework_14[Trework_14['DefectType']=='Rework']
Trework_14=Trework_14.groupby(['Date','BusinessUnit','DefectName']).agg({'DefectCount' :'sum','DefectType':'count'}).reset_index()
Trework_14['Rework']=Trework_14['DefectType']
del Trework_14['DefectType']
del Trework_14['DefectCount']
Trework_14

TallD_ok=pd.merge(TallD,TOk_count_14,on=['Date','BusinessUnit','DefectName'])
TallD_ok
TallD_ok_rej=pd.merge(TallD_ok,Treject_count_14,on=['Date','BusinessUnit','DefectName'],how='left')
TallD_ok_rej
TallD_ok_rej_rw=pd.merge(TallD_ok_rej,Trework_14,on=['Date','BusinessUnit','DefectName'],how='left')

TallD_ok_rej_rw=TallD_ok_rej_rw.fillna(0)

TallD_ok_rej_rw['NewTotalGarments']=TallD_ok_rej_rw['Total Garments']-(allD_ok_rej_rw['Reject']+allD_ok_rej_rw['Rework'])
TallD_ok_rej_rw=TallD_ok_rej_rw.fillna(0)
TallD_ok_rej_rw['Reject']=TallD_ok_rej_rw['Reject'].astype(int)
TallD_ok_rej_rw['Rework']=TallD_ok_rej_rw['Rework'].astype(int)
TallD_ok_rej_rw['NewTotalGarments']=TallD_ok_rej_rw['NewTotalGarments'].astype(int)
del TallD_ok_rej_rw['NewTotalGarments']
TallD_ok_rej_rw=TallD_ok_rej_rw[['Date','BusinessUnit','DefectName','Total Defect','Total Garments','Ok', 'Reject', 'Rework']]
TallD_ok_rej_rw

TallD_ok_rej_rw['Check Qty']=TallD_ok_rej_rw['Ok']
TallD_ok_rej_rw['Defect Qty']=TallD_ok_rej_rw['Total Defect']
del TallD_ok_rej_rw['Ok']
del TallD_ok_rej_rw['Total Defect']
TallD_ok_rej_rw['DHU%']=(TallD_ok_rej_rw['Defect Qty'])/TallD_ok_rej_rw['Total Garments']
TallD_ok_rej_rw['Reject%']=(TallD_ok_rej_rw['Reject']*100)/TallD_ok_rej_rw['Total Garments']
TallD_ok_rej_rw=TallD_ok_rej_rw.round(2)
TallD_ok_rej_rw

unit_w=TallD_ok_rej_rw.copy()

unit_1=unit_w[unit_w['BusinessUnit']=='Unit-1']
unit_1=unit_1[['BusinessUnit','DefectName', 'DHU%']].sort_values(['DHU%'], ascending=False).head(3)
unit_1


unit_2=unit_w[unit_w['BusinessUnit']=='Unit-2']
unit_2=unit_2[['BusinessUnit','DefectName', 'DHU%']].sort_values(['DHU%'], ascending=False).head(3)
unit_2

unit_3=unit_w[unit_w['BusinessUnit']=='Unit-3']
unit_3=unit_3[['BusinessUnit','DefectName','DHU%']].sort_values(['DHU%'], ascending=False).head(3)
unit_3

unit_4=unit_w[unit_w['BusinessUnit']=='Unit-4']
unit_4=unit_4[['BusinessUnit','DefectName','DHU%']].sort_values(['DHU%'], ascending=False).head(3)
unit_4

frames = [unit_1,unit_2,unit_3,unit_4]

result = pd.concat(frames)
result

'''############################ Run App #########################################################'''

server = Flask(__name__)
app = dash.Dash(server=server, external_stylesheets=[dbc.themes.SUPERHERO],
                meta_tags=[{'name': 'viewport',
                            'content': 'width=device-width, initial-scale=1.0, maximum-scale=1.2, minimum-scale=0.5,'}])

unit = mao2.BusinessUnit.unique().tolist()

app.layout = html.Div([
    dcc.Location(id='url', refresh=False),
    html.Div(id='page-content')
])

index_page = html.Div([
    dcc.Link('Report 14: Daily Sewing Quality Summary Report', href='/page-1'),
    html.Br(),
    dcc.Link('Page 2: Linewise Hourly Production Report', href='/page-2'),
    html.Br(),
    #     dcc.Link('Page 3: Daily Production Summary', href='/page-3'),
    #     html.Br(),
    #     dcc.Link('Page 4: New Report to Add ', href='/page-4'),
])
'''######################################### Page 1 ###############################'''

page_1_layout = html.Div([
    dbc.Row([
        dbc.Col([
            dash_table.DataTable(
                id='Unit-wise-Table',
                data=allD_ok_rej_rw.to_dict('records'),
                page_current=0,
                page_size=20,
                editable=False,
                columns=[{"name": i, "id": i, "deletable": True} for i in allD_ok_rej_rw.columns],
                style_cell_conditional=[
                    {
                        'font-size': '15px',
                        'color': 'Black',
                        'background-color': 'lightgrey',
                        'border': 'solid slategrey',
                        'border-width': '0.1px'
                    }
                ],
            ),
            html.Button("Download Unit Wise Concated CSV", id="btn_csv_1"),
            dcc.Download(id="download-unit-wise-concated-csv"),
        ], width={"size": 12, "offset": 0}, md=12),
    ], align="start"),

    dbc.Row([
        dbc.Col([
            dash_table.DataTable(
                id='Top-Defect-Table',
                data=result.to_dict('records'),
                page_current=0,
                page_size=20,
                editable=False,
                columns=[{"name": i, "id": i, "deletable": True} for i in result.columns],
                style_cell_conditional=[
                    {
                        'font-size': '15px',
                        'color': 'Black',
                        'background-color': 'lightgrey',
                        'border': 'solid slategrey',
                        'border-width': '0.1px'
                    }
                ],
            ),
            html.Button("Download Unit Wise Concated CSV", id="btn_csv_1"),
            dcc.Download(id="download-unit-wise-concated-csv"),
        ], width={"size": 12, "offset": 0}, md=12),
    ]),

    ##########################  PAGE  #########################
    html.Div(id='page-1-content'),
    html.Br(),
    dcc.Link('Linewise Hourly Production Report', href='/page-2'),
    html.Br(),
    dcc.Link('Go back to home', href='/'),
])


@app.callback(
    Output("download-unit-wise-concated-csv", "data"),
    Input("btn_csv_1", "n_clicks"),
    prevent_initial_call=True,
)
def func(n_clicks):
    return dcc.send_data_frame(Prod_Fore_plan_Unit.to_csv, "UnitWiseConcatedReport.csv")


'''######################################### Page 2 ###############################'''
page_2_layout = html.Div([
    dbc.Container([
        dbc.Row([
            dbc.Col(html.Div('Logo',
                             style={'text-align': 'center', 'color': "White",
                                    'height': '50px',
                                    'font-size': '30px',
                                    'font-family': 'Georgia',
                                    'background-color': 'Black',
                                    'border-style': 'Double',
                                    'border-color': 'Grey'}),
                    md=4),

            dbc.Col(html.Div('Floor Name',
                             style={'text-align': 'center', 'color': "White",
                                    'height': '50px',
                                    'font-size': '30px',
                                    'font-family': 'Georgia',
                                    'background-color': 'Black',
                                    'border-style': 'Double',
                                    'border-color': 'Grey'}
                             ), md=4),
            #                dbc.Col(
            #                 html.Div(datetime.datetime.now().strftime('%d-%m-%Y'),
            #                          style={'text-align':'center',
            #                                 'color':"White",
            #                                 'height':'50px',
            #                                 'font-size':'30px',
            #                                 'font-family':'Georgia',
            #                                 'background-color':'Black',
            #                                 'border-style':'Double',
            #                                 'border-color': 'Grey'}
            #                         ),md=4),

        ]),
        dbc.Row([
            html.H3("Choose Unit"),
            dcc.Dropdown(
                id='Unit_dropdown',
                options=[{'label': un, 'value': un} for un in unit],
                value="Unit",
                placeholder="Unit",
                style={'color': "Black", 'font-size': '25px', 'width': '50%'}
            ),
        ]),
    ]),
    html.Div(

        html.Div([
            dash_table.DataTable(
                id='Line-wise-table-cont',
                data=mao2.to_dict('records'),
                columns=[
                    {"name": i, "id": i, "deletable": False, "selectable": False} for i in mao2.columns
                ],
                page_size=15,
                editable=False,
                row_selectable="multi",
                row_deletable=False,
                selected_rows=[],
                page_action="native",

                style_cell_conditional=[
                    {
                        'color': "white",
                        'font-size': '20px',
                        'text-align': "center",
                        'background-color': 'black',
                        'border-style': 'double',
                        # 'display':'inline-block',
                    }
                ],
            ),
        ], className='row', style={'overflowY': 'scroll'}),
    ),

    html.Button("Download Line Wise CSV", id="btn_csv4"),
    dcc.Download(id="download-line-wise-csv"),
    ##########################  PAGE  #########################
    html.Div(id='page-2-content'),
    html.Br(),
    dcc.Link('Unitwise Hourly Production Report', href='/page-1'),
    html.Br(),
    dcc.Link('Go back to home', href='/'),
    html.Br(),
    dcc.Link('Daily Production Summary', href='/page-3'),
    html.Br(),
])


@app.callback(
    Output("download-line-wise-csv", "data"),
    Input("btn_csv4", "n_clicks"),
    prevent_initial_call=True,
)
def func(n_clicks):
    return dcc.send_data_frame(allD_ok_rej_rw.to_csv, "LineWiseReport.csv")


######################### Function ################################

@app.callback(
    Output('Line-wise-table-cont', 'data'),
    [Input('Unit_dropdown', 'value')],
)
def display_table(unit):
    Linefil = mao2[mao2.BusinessUnit == unit]
    return Linefil.to_dict('records')


###################################################################

@app.callback(Output('page-content', 'children'),
              [Input('url', 'pathname')])
def display_page(pathname):
    if pathname == '/page-1':
        return page_1_layout
    elif pathname == '/page-2':
        return page_2_layout
    #     elif pathname == '/page-3':
    #             return page_3_layout
    #     elif pathname == '/page-4':
    #         return page_4_layout
    else:
        return index_page


if __name__ == '__main__':
    app.run_server(debug=False)


import openpyxl
from openpyxl import Workbook, load_workbook
import pandas as pd
import csv

import time
from time import mktime
from datetime import datetime
# from datetime import date

import sys
import numpy as np

import matplotlib.pyplot as plt
import plotly
import chart_studio.plotly as py
import seaborn as sns
from plotly import graph_objs as go
import plotly.express as px
#import datatable as dt

import dash
import dash_table
from datetime import date
from dash import Dash, dcc, html
from dash.dependencies import Input,State, Output
# import dash_core_components as dcc
# import dash_html_components as html
import dash_bootstrap_components as dbc
# from dash.dependencies import Input, Output
# import vaex
#from jupyter_dash import JupyterDash  # pip install dash==1.19.0 or higher
from dash import callback_context
import pyodbc

import requests
import csv
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
       'DefectName', 'DefectCount']]

raw_api1

raw_api3=raw_api1.copy()
raw_api3=raw_api1.groupby(['BusinessUnit', 'LineNumber','Date','StyleSubCat','PoNumber',
                           'Color','DefectName']).agg({'Unique Garments':'count','DefectCount':'sum'}).reset_index()
raw_api3

raw_api4=raw_api1.copy()
raw_api4=raw_api4.groupby(['BusinessUnit', 'LineNumber','Date','StyleSubCat','PoNumber','Color','DefectName','Unique Garments']).agg({'DefectCount':'sum'}).reset_index()

#counting Non Defective garments
raw_api5=raw_api4.copy()
raw_api5=raw_api5[raw_api5['DefectName']=='na']
raw_api5=raw_api5.groupby(['BusinessUnit', 'LineNumber','Date','StyleSubCat','PoNumber','Color']).agg({'DefectCount':'sum','Unique Garments':'count'}).reset_index()
raw_api5['Non Defective Garments']=raw_api5['Unique Garments']
del raw_api5['Unique Garments']
del raw_api5['DefectCount']
raw_api5

#counting Defective garments
raw_api6=raw_api4.copy()
raw_api6=raw_api6[raw_api6['DefectName']!='na']
raw_api6=raw_api6.groupby(['BusinessUnit', 'LineNumber','Date','StyleSubCat','PoNumber','Color']).agg({'DefectCount':'sum','Unique Garments':'count'}).reset_index()
raw_api6['Defective Garments']=raw_api6['Unique Garments']
del raw_api6['Unique Garments']
raw_api6['Total Defects']=raw_api6['DefectCount']
del raw_api6['DefectCount']
raw_api6

# merge defect and non defect
merge1=pd.merge(raw_api5,raw_api6,on=['BusinessUnit', 'LineNumber','Date','StyleSubCat','PoNumber','Color'],how='outer')
merge1=merge1.fillna(0)
merge1['Check Qty']=merge1['Non Defective Garments']+ merge1['Defective Garments']
merge1['DHU%']=(merge1['Total Defects']*100)/merge1['Check Qty']
merge1

#Pivoting Defect Names
raw_api8=raw_api4[raw_api4['DefectName']!='na']
raw_api8=raw_api8.pivot_table('DefectCount', ['BusinessUnit', 'LineNumber','Date','StyleSubCat', 'PoNumber',
                           'Color'], 'DefectName').reset_index()
raw_api8=raw_api8.fillna(0)
raw_api8

# merging pivot table and defect/non defect
merge2=pd.merge(merge1,raw_api8,on=['BusinessUnit', 'LineNumber','Date','StyleSubCat','PoNumber','Color'],how='outer')
merge2=merge2.fillna(0)
# merge2.iloc[:,6:]=merge2.iloc[:,6:].astype(int)
merge2.iloc[:,6:10]=merge2.iloc[:,6:10].astype(int)
merge2.loc[:,'DHU%']=merge2.loc[:,'DHU%'].astype(float)
merge2.iloc[:,11:]=merge2.iloc[:,11:].astype(int)
merge2=merge2.round(2)
excel=merge2.copy()
del excel['Non Defective Garments']
excel

# # excel1=excel.to_excel('C:\\Users\\nadiajebin\\Desktop\\send\\Defect Report.xlsx')

# DashBoard Line and Unit wise
pew = excel[['Date','BusinessUnit','LineNumber','StyleSubCat','Color','Check Qty','Total Defects','Defective Garments']].copy()
pew = pew.groupby(['Date','BusinessUnit','LineNumber']).agg(
    {'Check Qty': 'sum', 'Total Defects': 'sum', 'Defective Garments': 'sum'}).reset_index()

dwgc = raw_api1.copy()
dwgc = dwgc.groupby(['Date','BusinessUnit','LineNumber']).agg({'GarmentsNumber': 'nunique', 'DefectCount': 'sum'}).reset_index()

# Defect postion data
defectpost = raw_api.copy()
defectpost = defectpost.groupby(['Date','BusinessUnit', 'LineNumber', 'StyleSubCat','Color', 'DefectName', 'DefectPos']).agg(
    {'GarmentsNumber': 'nunique', 'DefectCount': 'sum'}).reset_index()
filt_defect = defectpost.copy()
filt_defectname = defectpost.copy()
filt_defect = filt_defect.groupby(['Date','BusinessUnit','LineNumber', 'DefectPos']).agg({'DefectCount': 'sum'}).reset_index()
filt_defectname = filt_defectname.groupby(['Date','BusinessUnit','LineNumber', 'DefectName']).agg(
    {'GarmentsNumber': 'sum'}).reset_index()
filt_defect = filt_defect[filt_defect['DefectCount'] > 0]
filt_defectname = filt_defectname[filt_defectname['DefectName'] != 'na']

# Plot
plt_hour = raw_api.copy()
# plt_hour['Time'] = plt_hour['Time'].apply(lambda t: t.strftime('%H'))
# plt_hour['Time'] = plt_hour['Time'].astype(str)
# plt_hour['Time'] = plt_hour['Time'] + ":00"
plt_hour = plt_hour.groupby(['Date', 'BusinessUnit', 'LineNumber', 'Time']).agg(
    {'GarmentsNumber': 'nunique', 'DefectCount': 'sum'}).reset_index()
# plt_hour = plt_hour.rename(columns={'Date': 'Date'})
plt_hour = plt_hour.sort_values(['Date', 'BusinessUnit', 'LineNumber', 'Time', 'GarmentsNumber'], ascending=True)

plt_hour_f = plt_hour.copy()
plt_hour_f = plt_hour_f.groupby(['Time']).agg({'GarmentsNumber': 'sum'}).reset_index()
plt_hour_f = plt_hour_f.sort_values(['Time', 'GarmentsNumber'], ascending=True)

################### Only Defect Data ###############
defect_only=excel[excel['Defective Garments']>0]


# App Start
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.SUPERHERO])
server=app.server

date  = defect_only.Date.unique().tolist()
unit  = defect_only.BusinessUnit.unique().tolist()
line  = defect_only.LineNumber.unique().tolist()
style = defect_only.StyleSubCat.unique().tolist()

app.layout = html.Div([
    dcc.Location(id='url', refresh=False),
    html.Div(id='page-content')
])

index_page = html.Div([
    dcc.Link('Page 1: Filter Table', href='/page-1'),
    html.Br(),
    dcc.Link('Page 2: Summary Table', href='/page-2'),
    html.Br(),
    dcc.Link('Page 3: Dash', href='/page-3'),
    html.Br(),
    dcc.Link('Page 4: Drag & Drop Hourly Report', href='/page-4'),
])


'''######################################### Page 1 ###############################'''
page_1_layout = html.Div([
    html.H1('Page 1: Filter Data'),
    html.Div(

        children=[
            html.Div('Filtered Dash Board',
                     style={
                         'height': '50px',
                         'font-size': '25px',
                         'font-family': 'Georgia',
                         'text-align': "left",
                         'display': 'block',
                         'width': '22%',
                     }
                     ),
            dcc.Dropdown(
                id='Date_dropdown',
                options=[{'label': st, 'value': st} for st in date],
                value="Date",
                placeholder="Date",
                style={'color': "Black", 'background-color': 'darkcyan', 'font-size': '25px', 'display': 'inline-block',
                       'width': '50%'}
            ),  # 'border-radius':'4px'

            dcc.Dropdown(
                id='Unit_dropdown',
                options=[{'label': un, 'value': un} for un in unit],
                value="Unit",
                placeholder="Unit",
                style={'color': "Black", 'background-color': 'darkcyan', 'font-size': '25px', 'display': 'inline-block',
                       'width': '50%'}
            ),

            dcc.Dropdown(
                id='Style_dropdown',
                options=[{'label': sty, 'value': sty} for sty in style],
                value="Style",
                placeholder="Style",
                style={'color': "Black", 'background-color': 'darkcyan', 'font-size': '15px', 'display': 'inline-block',
                       'width': '50%'}
            ),

            dcc.Dropdown(
                id='Line_dropdown',
                options=[{'label': ln, 'value': ln} for ln in line],
                value="Line",
                placeholder="Line",
                style={'color': "Black", 'background-color': 'darkcyan', 'font-size': '25px', 'display': 'inline-block',
                       'width': '50%'}
            ),

            dash_table.DataTable(
                id='table-container',
                data=defect_only.to_dict('records'),
                style_cell_conditional=[
                    {
                        'color': "Black",
                        'font-size': '20px',
                        'text-align': "center",
                        'background-color': 'darkcyan',
                        'border-style': 'solid',
                    }
                ],
            )
        ],

    ),

    html.Div(id='page-1-content'),
    html.Br(),
    dcc.Link('Go to Page 2', href='/page-2'),
    html.Br(),
    dcc.Link('Go back to home', href='/'),
])


'''######################################### Page 2 ###############################'''
page_2_layout = html.Div([
    html.H1('Page 2: Summary Table'),
    html.Div('Dash Board',
             style={
                 'color': "White",
                 'height': '50px',
                 'font-size': '30px',
                 'font-family': 'Georgia',
                 'text-align': "center",
                 'background-color': 'Black',
                 'border-style': 'Double',
                 'border-color': 'Grey',
                 'display': 'block',
                 'width': '20%',
             }
             ),
    html.Div(

        html.Div([
            dash_table.DataTable(
                id='datatable_id',
                data=excel.to_dict('records'),
                columns=[
                    {"name": i, "id": i, "deletable": False, "selectable": False} for i in excel.columns
                ],
                editable=False,
                filter_action='native',
                sort_action="native",
                sort_mode="multi",
                row_selectable="multi",
                row_deletable=False,
                selected_rows=[],
                page_action="native",
                page_current=0,
                page_size=20,

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
        ], className='row'), ),
    ###################################################
    html.Div(id='page-2-content'),
    html.Br(),
    dcc.Link('Go to Filter Page', href='/page-1'),
    html.Br(),
    dcc.Link('Go back to home', href='/'),
    html.Br(),
    dcc.Link('Go to Dash', href='/page-29'),
    html.Br(),
])

'''######################################### Page 03 ###############################'''

page_3_layout = html.Div([
    html.H1('Defect Report: '),
    dbc.Container([
        dbc.Row([
        ]),
        dbc.Row([
            dbc.Col([
                dbc.Row([
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.H6("Insert Values"),
                                dbc.Row([
                                    dbc.Col([
                                        dbc.Input(id='input1', value='Start Date', type='text',
                                                  style={'display': 'inline-block'}), ]),
                                    dbc.Col([
                                        dbc.Input(id='input2', value='End date', type='text',
                                                  style={'display': 'inline-block'}), ]), ]),
                                dbc.Row([
                                    dbc.Col([
                                        dbc.Input(id='unit_input', value='Unit', type='text', style={"width": '1'}), ]),
                                    dbc.Col([
                                        dbc.Input(id='line_input', value='Line', type='number',
                                                  style={'display': 'inline-block'}), ]), ]),
                                dbc.Button(id='submit-button1', type='submit', children='Submit', size="sm"),

                            ])
                        ])
                    ], width=12),
                ]),

                dbc.Row(
                    dbc.Col(html.Hr(style={'border': "3px solid gray"}), width=12)
                ),
                dbc.Row([
                    dbc.Col(dbc.Card([
                        dbc.CardBody([
                            html.H6("Total Production"),
                            html.H4(id="totalProduction", style={'fontWeight': 'bold'}),
                        ])
                    ]), width=12)
                    , ], className="mb-3"),

                dbc.Row([
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.H6("Defects"),
                                html.H4(id="totaldefect", style={'fontWeight': 'bold'}),

                            ])
                        ])
                    ], width=6),
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.H6("Defective Garments"),
                                html.H4(id="UniqueDefectCount", style={'fontWeight': 'bold'}),
                            ])
                        ])
                    ], width=6),

                ], className="mb-3"),

                dbc.Row([
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.H6("DHU%"),
                                html.H4(id="dhu", style={'fontWeight': 'bold'}),

                            ])
                        ])
                    ], width=6),
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.H6("Defect%"),
                                html.H4(id='output_div', style={'fontWeight': 'bold'}),
                            ])
                        ])
                    ], width=6),
                ], className="mb-3"),

                dbc.Row([
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                dcc.Link('Home', href='/'),
                                dcc.Link('Data Table', href='/page-2', style={"margin-left": "10px"}),
                                dcc.Link('Filter ', href='/page-1', style={"margin-left": "10px"}),
                                dcc.Link('Drag & Drop', href='/page-4', style={"margin-left": "10px"}),

                            ])
                        ])
                    ], width=12), ]),

            ], width=4),

            dbc.Col([
                dbc.Row([
                    dbc.Card([
                        dbc.CardBody([
                            html.P("No of Unique Garment by Hour"),
                            dcc.Graph(id="bar-chart", config={'displayModeBar': True},
                                      figure=px.bar(plt_hour_f, x='Time', y='GarmentsNumber',
                                                    text="GarmentsNumber").
                                      update_layout(autosize=False, width=700, height=250),
                                      )
                        ])
                    ])
                ]),
                dbc.Row(
                    dbc.Col(html.Hr(style={'border': "3px solid gray"}), width=12)
                ),
                dbc.Row([
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.H3(id='Datatablewithtable',
                                        children=html.Div([
                                            dash_table.DataTable(
                                                id='table-dropdown',
                                                data=filt_defectname.to_dict('records'),
                                                page_current=0,
                                                page_size=3,
                                                editable=False,
                                                columns=[{"name": i, "id": i, "deletable": True} for i in
                                                         filt_defectname.iloc[:, 3:]],
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
                                        ], style={'fontWeight': 'bold', 'textAlign': 'center', 'color': 'white'})
                                        ),
                            ]), ]), ], width=6, md=6),

                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.H3(id='Datatablewithtable2',
                                        children=html.Div([
                                            dash_table.DataTable(
                                                id='table-dropdown2',
                                                data=filt_defect.to_dict('records'),
                                                page_current=0,
                                                page_size=3,
                                                editable=False,
                                                columns=[{"name": i, "id": i, "deletable": True} for i in
                                                         filt_defect.iloc[:, 3:]],
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
                                        ], style={'fontWeight': 'bold', 'textAlign': 'center', 'color': 'white'}),
                                        )  #

                            ])
                        ]), ], width=6, md=6),
                ])
            ], width=7)
        ], className="mt-3")
    ], fluid=True, style={'backgroundColor': 'lightgrey'}),

])


'''######################################### Page 4 ###############################'''

page_4_layout = html.Div([
        dcc.Upload(
          id='upload-data',
          children=html.Div([
            'Drag and Drop or ',
            html.A('Select Files')
          ]),
          style={
            'width': '100%',
            'height': '60px',
            'lineHeight': '60px',
            'borderWidth': '1px',
            'borderStyle': 'dashed',
            'borderRadius': '5px',
            'textAlign': 'center',
            'margin': '10px'
           },
          # Allow multiple files to be uploaded
          multiple=True
         ),
       dbc.Container([
        dbc.Row([
                dbc.Col(html.Div("LOGO", style={'text-align':'left'}), md=4),
                dbc.Col(html.Div("Galileo", style={'text-align':'center'}), md=4),
                dbc.Col(html.Div(datetime.now().strftime('%d-%m-%Y'), style={'text-align':'right'}), md=4),

        ],justify="between"),
        dbc.Row([
            dbc.Col([
                dbc.Row([
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.Div(id='output-datatable' ),
                                dbc.Button(id='btn',
                                           children=[html.I(className="fa fa-download mr-1"), "Download"],
                                           color="info",
                                           className="mt-1"
                                           ),


                        ])
                    ], ),
                ],width=12),

                dbc.Row(
                     dbc.Col(html.Hr(style={'border': "3px solid gray"}), width=12)
                    # html.Div(id='output-datatable'),
                ),
                dbc.Row(


                ),
                dbc.Row([
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.Div("Sign Off Plan : 0"),
                            ])
                        ])
                    ], width=2),
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.Div("Day Plan Target : 14030"),
                            ])
                        ])
                    ], width=3),
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.Div("Actual PCS : 3869"),
                            ])
                        ])
                    ], width=2),
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.Div("PLAN Eff% : 2.16"),
                            ])
                        ])
                    ], width=2),
                    dbc.Col([
                        dbc.Card([
                            dbc.CardBody([
                                html.Div("Actual Eff% : 0.53", style={'text-align':'right'}),
                            ])
                        ])
                    ], width=2),
                ]),
                    dbc.Row(html.Div(id='output-div'),),
            ], ),

            # dbc.Col([
            #
            # ], width=1)
        ], className="mt-3")
    ]),

    ], fluid=True,),
    html.Div(id='page-4-content'),
    html.Br(),
    dcc.Link('Go back to home', href='/'),
    html.Br(),
    dcc.Link('Go to Dash', href='/page-3'),
    html.Br(),
    ])

''' ##################################### FUNCTIONS ##########################################'''


############################# Page-1 Dash Filter Table Starts ######################
@app.callback(
    Output('table-container', 'data'),
    [Input('Date_dropdown', 'value')],
    [Input('Unit_dropdown', 'value')],
    [Input('Line_dropdown', 'value')],
    [Input('Style_dropdown', 'value')],
)
def display_table(date, unit, line, style):
    df1 = defect_only[defect_only.Date == date]
    df2 = df1[df1.BusinessUnit == unit]
    df3 = df2[df2.LineNumber == line]
    df4 = df3[df3.StyleSubCat == style]
    return df4.to_dict('records')


############################# Page-1 Dash Filter Table Ends ######################

############################# Page-3 Dash Functions Start ######################
@app.callback(
    Output("dhu", "children"),
    Output("totaldefect", "children"),
    Output("totalProduction", "children"),
    Output('UniqueDefectCount', "children"),
    Output('output_div', 'children'),
    Output('bar-chart', 'figure'),
    Output('table-dropdown', 'data'),
    Output('table-dropdown2', 'data'),
    [Input('submit-button1', 'n_clicks')],
    [State('input1', 'value')],
    [State('input2', 'value')],
    [State('unit_input', 'value')],
    [State('line_input', 'value')],
)
def update_o(clicks, input1, input2, unit_input, line_input):
    # Static Output
    new2 = excel.loc[
        (excel['BusinessUnit'] == unit_input) & (excel['LineNumber'] == line_input), ['Date', 'BusinessUnit',
                                                                                      'LineNumber',
                                                                                      'Check Qty', 'Total Defects',
                                                                                      'Defective Garments']]
    start_date = new2.index[(new2["Date"] == input1)][0]
    end_date = new2.index[(new2["Date"] == input2)][0]

    # Bar Output
    newplt = plt_hour.loc[
        (plt_hour['BusinessUnit'] == unit_input) & (plt_hour["LineNumber"] == line_input), ['Date', 'BusinessUnit',
                                                                                            'LineNumber', 'Time',
                                                                                            'GarmentsNumber',
                                                                                            'DefectCount']]
    start_date1 = newplt.index[(newplt["Date"] == input1)].min()
    end_date1 = newplt.index[(newplt["Date"] == input2)].max()
    dff = newplt.loc[start_date1:end_date1, :]

    # Defect Name Table
    testu = filt_defectname.loc[
        (filt_defectname['BusinessUnit'] == unit_input) & (filt_defectname["LineNumber"] == line_input),
        ['Date', 'BusinessUnit', "LineNumber", 'DefectName', 'GarmentsNumber']]
    testu1 = testu.index[(testu["Date"] == input1)].min()
    testu2 = testu.index[(testu["Date"] == input2)].max()
    testudff = testu.loc[testu1:testu2, 'DefectName':]
    data = testudff.to_dict('records')
    columns = [{"name": i, "id": i, } for i in (testudff.columns)]

    # Defect Position Table
    testu_dp = filt_defect.loc[(filt_defect['BusinessUnit'] == unit_input) & (filt_defect["LineNumber"] == line_input),
                               ['Date', 'BusinessUnit', "LineNumber", 'DefectPos', 'DefectCount']]
    testu_st = testu_dp.index[(testu_dp["Date"] == input1)].min()
    testu_ed = testu_dp.index[(testu_dp["Date"] == input2)].max()
    testud_defp = testu_dp.loc[testu_st:testu_ed, 'DefectPos':]
    data2 = testud_defp.to_dict('records')
    columns = [{"name": i, "id": i, } for i in (testud_defp.columns)]

    if clicks is not None:
        GarmentsNumber = new2.loc[start_date:end_date, 'Check Qty'].sum()
        DefectCount = new2.loc[start_date:end_date, 'Total Defects'].sum()
        UniqueDefectCount = new2.loc[start_date:end_date, 'Defective Garments'].sum()
        DHU = (DefectCount * 100) / GarmentsNumber
        DHU_format = round(DHU, 2)
        DefectPercent = (UniqueDefectCount * 100) / GarmentsNumber
        DefectPercent_format = round(DefectPercent, 2)
        figure = px.bar(dff, x='Time', y='GarmentsNumber', text='GarmentsNumber', color='Date')

        return DHU_format, DefectCount, GarmentsNumber, UniqueDefectCount, DefectPercent_format, figure, data, data2


############################# Page-3 Dash Functions Ends ######################

############################# Page-4 Upload Page Functions Start ######################
@app.callback(Output('output-datatable', 'children'),
              Input('upload-data', 'contents'),
              State('upload-data', 'filename'),
              State('upload-data', 'last_modified'))
def update_output(list_of_contents, list_of_names, list_of_dates):
    if list_of_contents is not None:
        children = [
            parse_contents(c, n, d) for c, n, d in
            zip(list_of_contents, list_of_names, list_of_dates)]
        return children


@app.callback(Output('output-div', 'children'),
              Input('submit-button', 'n_clicks'),
              State('stored-data', 'data'),
              State('xaxis-data', 'value'),
              State('yaxis-data', 'value'))
def make_graphs(n, data, x_data, y_data):
    if n is None:
        return dash.no_update
    else:
        bar_fig = px.bar(data, x=x_data, y=y_data)
        # print(data)
        return dcc.Graph(figure=bar_fig)


@app.callback(
    Output("download-component", "data"),
    Input("btn", "n_clicks"),
    State('stored-data', 'data'),
    prevent_initial_call=True,
)
def func(n_clicks, data):
    empty = pd.DataFrame(data)
    return dcc.send_data_frame(empty.to_excel, "mydf_excel.xlsx", sheet_name="Sheet_name_1")


def parse_contents(contents, filename, date):
    content_type, content_string = contents.split(',')

    decoded = base64.b64decode(content_string)
    try:
        if 'csv' in filename:
            # Assume that the user uploaded a CSV file
            df = pd.read_csv(
                io.StringIO(decoded.decode('utf-8')))
        elif 'xls' in filename:
            # Assume that the user uploaded an excel file
            df = pd.read_excel(io.BytesIO(decoded))
            empty = df.copy()

    except Exception as e:
        print(e)
        return html.Div([
            'There was an error processing this file.'
        ])

    return html.Div([
        html.H6(filename, style={'display': 'inline-block'}),
        html.H6(datetime.datetime.fromtimestamp(date), style={'display': 'inline-block'}),
        html.P("Inset X axis data"),
        dcc.Dropdown(id='xaxis-data',
                     options=[{'label': x, 'value': x} for x in df.columns],
                     style={'color': "Black", 'background-color': 'White', 'font-size': '15px'}),
        html.P("Inset Y axis data"),
        dcc.Dropdown(id='yaxis-data',
                     options=[{'label': x, 'value': x} for x in df.columns],
                     style={'color': "Black", 'background-color': 'White', 'font-size': '15px'}),
        html.Button(id="submit-button", children="Create Graph"),
        html.Hr(),

        dash_table.DataTable(
            # empty=df.copy(),
            data=df.to_dict('records'),
            columns=[{'name': i, 'id': i} for i in df.columns],
            page_size=15,
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

        dcc.Store(id='stored-data', data=df.to_dict('records')),
        dcc.Download(id="download-component"),
        html.Hr(),  # horizontal line

    ])


############################# Page-4 Upload Page Functions Ends ######################

############################# Page CallBacks ######################
@app.callback(Output('page-content', 'children'),
              [Input('url', 'pathname')])
def display_page(pathname):
    if pathname == '/page-1':
        return page_1_layout
    elif pathname == '/page-2':
        return page_2_layout
    elif pathname == '/page-3':
        return page_3_layout
    elif pathname == '/page-4':
        return page_4_layout
    #     elif pathname == '/page-30':
    #         return page_30_layout
    else:
        return index_page
    # You could also return a 404 "URL not found" page here


if __name__ == '__main__':
    app.run_server(debug=False)


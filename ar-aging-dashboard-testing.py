import dash
from dash import Dash, dcc, html, Input, Output    # pip install dash
import dash_bootstrap_components as dbc       # pip install dash-bootstrap-components
import plotly.express as px                   # pip install plotly.express 
import dash_mantine_components as dmc
import pandas as pd
import datetime
import plotly.graph_objects as go
import numpy as np
from plotly.subplots import make_subplots


app = Dash(__name__, external_stylesheets = [dbc.themes.SPACELAB], meta_tags = [{ 'name': 'viewport', 'content': 'width = device-width, initial-scale = 1, maximum-scale = 1'}])

#-------------------------------------------------------------------------------------------------------------------------------------


# Importing Data and Data Clean Up

# -- Collections Data --

df_collections = pd.read_excel(r'C:\Users\massimo.biagiotti\Desktop\Cloud Services - Weekly AR Aging Input.xlsx', sheet_name = 'Collections - Cybersecurity')
df_collections.sort_values(by = 'Placement Date')
df_collections = df_collections.tail(13)
df_collections['Placement Date'] = pd.to_datetime(df_collections['Placement Date'], format = '%y%m%d')

# -- Top 5 Data --

df_top5 = pd.read_excel(r'C:\Users\massimo.biagiotti\Desktop\Cloud Services - Weekly AR Aging Input.xlsx', sheet_name = 'Top 5 Customer Input USD', skiprows = 3, usecols = np.r_[1:22])
df_top5 = df_top5.dropna(how = 'all')
df_top5 = df_top5.sort_values(by = '90+')
df_top5 = df_top5.rename(columns={'W/E Date': 'Month', 'BU Name': 'BU', 'Brand Master Group': 'Brand Master Grouping', 'Region Group': 'Location', '90+ % By BU': '90+ % by BU', 'Top 5  Customers': 'Customer' })


# -- AR Current Month Aging (Weekly View) --

df_cm_aging_weekly = pd.read_excel(r'C:\Users\massimo.biagiotti\Desktop\Cloud Services - Weekly AR Aging Input.xlsx', sheet_name = 'Aging USD', skiprows = 3, usecols = np.r_[0:17,19:25])

# dropping any null values from the dataframe
df_cm_aging_weekly = df_cm_aging_weekly.dropna(how = 'all')

# -- AR Aging Historical Archive (Monthly) --

df_aging_monthly = pd.read_excel(r'C:\Users\massimo.biagiotti\Desktop\Cloud Services - Weekly AR Aging Input.xlsx', sheet_name = 'Aging Monthly Historical Archiv')

# dropping any null values from the dataframe
df_aging_monthly = df_aging_monthly.dropna(how = 'all')


# Need to create a new Dataframe from df_cm_aging_weekly
# This Datafram will be appended to df_aging_monthly in order to create a historical archive along with a dynamic current month graph

# Cleaning and reorganizing the weekly data prior to appending
df_cm_aging_current_week_append = df_cm_aging_weekly
df_cm_aging_current_week_append = df_cm_aging_current_week_append.drop(columns=['Input Owner'])
df_cm_aging_current_week_append = df_cm_aging_current_week_append[['W/E Date', 'Brand', 'Brand Location Group', 'Brand Master Group', 'Consolidated BU','BU Name','Region Group', 'Country', 'Currency', 'Current', '1-30', '31-60', '61-90', '90+', 'Total A/R', '90+ %', '90+ % By BU', '90+ % Consolidated', '90+ % By Brand Location Group', '90+ % By Region Group', '90+ % By Brand Master Group', '90+ % By Consolidated BU']]

# Changing all weekly dates to first day of the month
df_cm_aging_current_week_append['W/E Date'] = pd.to_datetime(pd.DataFrame({'day': 1,
                                                                           'month': df_cm_aging_current_week_append['W/E Date'].dt.month,
                                                                           'year': df_cm_aging_current_week_append['W/E Date'].dt.year},
                                                                           index = df_cm_aging_current_week_append.index))
# Renaming columns to match Monthly columns    
df_cm_aging_current_week_append = df_cm_aging_current_week_append.rename(columns={'W/E Date': 'Month', 'BU Name': 'BU', 'Brand Master Group': 'Brand Master Grouping', 'Region Group': 'Location', '90+ % By BU': '90+ % by BU' })


# Appending df_cm_aging_current_week_append to df_aging_monthly

frames = [df_aging_monthly,df_cm_aging_current_week_append]
df_aging_monthly = pd.concat(frames)



# Creating DataFrame for Graphs

df_aging_monthly_consolidated = df_aging_monthly.groupby(['Month', 'Consolidated BU']).sum()

df_aging_monthly_consolidated.reset_index(inplace=True)

df_aging_monthly_region = df_aging_monthly.groupby(['Month', 'Consolidated BU' ,'Location']).sum()

df_aging_monthly_region.reset_index(inplace = True)

df_top5_region = df_top5
df_top5_region = df_top5.groupby(['BU','Location', 'Brand']).sum()
df_top5_region.reset_index(inplace = True)
df_top5_region = df_top5_region[['BU','Location','Brand', 'Customer', 'Current', '1-30', '31-60', '61-90', '90+', 'Total A/R', '90+ %']]


print(df_top5_region)


# print(df_aging_monthly_consolidated)
# print(df_aging_monthly_region)








#----------------------------------------------------------------------------------------------------------------------------------------

#Data testing
# print(df_collections)
# print(df_collections.dtypes)

# print(df_cm_aging_weekly)


# print(df_aging_monthly)
# print(df_aging_monthly.dtypes)

# print(df_cm_aging_current_week_append)
# print(df_cm_aging_current_week_append.dtypes)

#----------------------------------------------------------------------------------------------------------------------------------------

# Creating variables

# -- Collection Variables ---
cm_collections = df_collections.iloc[-1]['Outstanding Balance at 3rd Party Collections']
cm_collections_int = int(round(cm_collections/1000,0))

cm_total_placements = df_collections.iloc[-1]['Total']
cm_total_placements_int = int(round(cm_total_placements/1000,0))

cm_month_date = df_collections.iloc[-1]['Placement Date']
cm_month_name = cm_month_date.strftime("%B")


# cm_week_date = df_cm_aging_weekly['W/E Date'] = pd.to_datetime(pd.DataFrame({'day': df_cm_aging_weekly['W/E Date'].dt.day,
#                                                                            'month': df_cm_aging_weekly['W/E Date'].dt.month,
#                                                                            'year': df_cm_aging_weekly['W/E Date'].dt.year},
#                                                                            index = df_cm_aging_weekly.index))
cm_week_name = df_cm_aging_weekly.iloc[-1]['W/E Date']
cm_week_name = cm_week_name.strftime("%b-%d-%y")

print(cm_week_name)
#Aging Variables ---


cm_aging_month = df_aging_monthly_consolidated.iloc[-1]['Month']
pm_aging_month = df_aging_monthly_consolidated.iloc[-2]['Month']

# print(cm_aging_month, pm_aging_month)

# cybersecurity_cm_aging_90amt = df_aging_monthly_consolidated.loc[(df_aging_monthly_consolidated['Consolidated BU'] == 'CyberSecurity') & (df_aging_monthly_consolidated['Month'] == cm_aging_month)]
# cybersecurity_pm_aging_90amt = df_aging_monthly_consolidated.loc[(df_aging_monthly_consolidated['Consolidated BU'] == 'CyberSecurity') & (df_aging_monthly_consolidated['Month'] == pm_aging_month)]

# cybersecurity_na_cm_aging_90amt

# print(cm_aging_90amt, pm_aging_90amt)

# ar_cm_90pct_bu = df_aging_monthly_consolidated.iloc[-1]['90+ % by BU']
# ar_cm_90pct_region = df_aging_monthly_region.iloc[-1]['90+ % By Region Group']



#----------------------------------------------------------------------------------------------------------------------------------------
#Data testing
# print(cm_date)
# print(cm_month_name)

# print(df_aging_monthly_region)

# cm_aging_90amt = df_aging_monthly_region.loc[(df_aging_monthly_region['Consolidated BU'] == 'CyberSecurity') & (df_aging_monthly_region['Month'] == cm_aging_month) & (df_aging_monthly_region['Location'] == 'North America')]

# print(cm_aging_90amt)

# cm_aging_90amt = cm_aging_90amt.iloc[-1]['90+']

# print(cm_aging_90amt)

#----------------------------------------------------------------------------------------------------------------------------------------
#Building Components


#----------------------------------------------------------------------------------------------------------------------------------------
#Creating Graphs

collections_gragh = go.Figure()

collections_gragh.add_trace(
    go.Scatter(
        x = df_collections['Placement Date'],
        y = df_collections['Total'],
        name = 'Total Placements',
    ))

collections_gragh.add_trace(
    go.Bar(
        x = df_collections['Placement Date'],
        y = df_collections['Outstanding Balance at 3rd Party Collections'],
        name = 'Outstanding Balance at 3rd Party Collections'
    ))

collections_gragh.update_layout(
    title = ("Outstanding Balance at 3rd Party Collections = " +'$'+str(cm_collections_int) +'k' + " " + " & " + "Total Placements = "+"$" +str(cm_total_placements_int)+'k' + " " + "for" + " " + cm_month_name)

    )







region_list = [{'label': i, 'value': i} for i in df_aging_monthly_region['Location'].unique()]
region_list.insert(0,{'label': 'Consolidated', 'value': 'Consolidated'})


#---------------------------------------------------------------
#App layout

app.layout = html.Div([

   html.H1("AR Aging Dashboard", style = {'text-align': 'center'}),

   html.Br(),

   html.H2("Weekly AR Report (in $1,000 USD) - " + str(cm_week_name), style = {'text-align': 'center'}),

   html.Br(),

   html.Br(),

   dcc.Dropdown(id = "select_bu",
                options =[
                {'label': i, 'value':i} for i in df_aging_monthly_consolidated['Consolidated BU'].unique()],
                     multi = False,
                     clearable = False,
                     value = 'CyberSecurity',
                     placeholder = 'Select Business Unit',
                     style = {'width': '30%'}
                     ),

   dcc.Dropdown(id = 'select_region',
                options = region_list,
                     multi = False,
                     clearable = False,
                     value = 'North America',
                     style = {'width': '30%'}
                     ),

   dcc.Graph(id = 'ar_aging', figure = {}),

   dcc.Graph(id = 'ar_aging_indicators', figure = {}),

   dcc.Graph(id = 'top_5', figure = {}),
 
   dcc.Graph(id = 'total_collections', figure = collections_gragh),

   


])


# -----------------------------------------------------------------
# Callback allows components to interact

# @app.callback(
#     [Output(component_id = 'select_region', component_property = 'options')],
#     [Output(component_id = 'select_region', component_property = 'value')],
#     [Input( component_id = 'select_bu', component_property = 'value')]

# )

# def update_dropdown(region_slctd):

#     df_aging_monthly_region_test = df_aging_monthly_region 
#     df_aging_monthly_region_test = df_aging_monthly_region_test[df_aging_monthly_region_test['Location'] == region_slctd]

#     region_list = [{'label': i, 'value': i} for i in df_aging_monthly_region['Location'].unique()]
#     region_list.insert(0,{'label': 'Consolidated', 'value': 'Consolidated'})

#     values_selected =  [x['value'] for x in region_list]

#     return region_list, values_selected




@app.callback(
    [Output(component_id = 'ar_aging', component_property = 'figure')],
    [Output(component_id = 'ar_aging_indicators', component_property = 'figure')],
    [Output(component_id = 'top_5', component_property = 'figure')],
    [Input(component_id = 'select_bu', component_property = 'value')],
    [Input(component_id = 'select_region', component_property = 'value')]

)

def update_graph(option_slctd, region_slctd):
    print(option_slctd)
    print(type(option_slctd))

     


    df_aging_monthly_consolidated_test = df_aging_monthly_consolidated
    df_aging_monthly_consolidated_test = df_aging_monthly_consolidated_test[df_aging_monthly_consolidated_test['Consolidated BU'] == option_slctd]

    df_aging_monthly_region_test = df_aging_monthly_region 
    df_aging_monthly_region_test = df_aging_monthly_region_test[df_aging_monthly_region_test['Location'] == region_slctd]
    df_aging_monthly_region_test = df_aging_monthly_region_test[df_aging_monthly_region_test['Consolidated BU'] == option_slctd]

    df_top5_test = df_top5_region
    df_top5_test = df_top5_test[df_top5_test['Location'] == region_slctd]
    df_top5_test = df_top5_test[df_top5_test['BU'] == option_slctd]



    ar_aging_title1 = (option_slctd + ' ' + region_slctd + ' ' + 'AR Aging' )

    # cm_aging_90amt = df_aging_monthly_region.loc[(df_aging_monthly_region['Consolidated BU'] == option_slctd) & (df_aging_monthly_region['Month'] == cm_aging_month) & (df_aging_monthly_region['Location'] == region_slctd)]

    # cm_aging_90amt = cm_aging_90amt.iloc[-1]['90+']




    #ar_aging_title2 = ('<span style="font-size: 21px;">Percentage of 90+ Over Total AR = ' + str(round(df_aging_monthly_consolidated_test.iloc[-1]['90+ % by BU'].sum()*100,1)) + '%' +'</span>')
    #ar_aging_title3 = ('<span style="font-size: 21px;">Percentage of 90+ Over Total AR = ' + str(round(df_aging_monthly_region_test.iloc[-1]['90+ % By Region Group'].sum()*100,1)) + '%' +'</span>')



    # ar_aging_title2 = ('<span style="font-size: 21px;">Percentage of 90+ Over Total AR = ' + str(round(ar_cm_90pct_bu*100,1)) + '%' +'</span>')
    # ar_aging_title3 = ('<span style="font-size: 21px;">Percentage of 90+ Over Total AR = ' + str(round(ar_cm_90pct_region*100,1)) + '%' +'</span>')



    if region_slctd == 'Consolidated':

        try:
            cm_aging_90amt = df_aging_monthly_consolidated_test.loc[(df_aging_monthly_consolidated_test['Consolidated BU'] == option_slctd) & (df_aging_monthly_consolidated_test['Month'] == cm_aging_month)]
            cm_aging_90amt = cm_aging_90amt.iloc[-1]['90+']

            pm_aging_90amt = df_aging_monthly_consolidated_test.loc[(df_aging_monthly_consolidated_test['Consolidated BU'] == option_slctd) & (df_aging_monthly_consolidated_test['Month'] == pm_aging_month)]
            pm_aging_90amt = pm_aging_90amt.iloc[-1]['90+']

        
        except:
            cm_aging_90amt = 0
            pm_aging_90amt = 0




        try:
            ar_aging_title2 = ('<span style="font-size: 21px;">Percentage of 90+ Over Total AR = ' + str(round(df_aging_monthly_consolidated_test.iloc[-1]['90+ % by BU'].sum()*100,1)) + '%' +'</span>')
        except BaseException:
            ar_aging_title2 = 'No Data Present'
       

        

        

        aging_indicators = go.Figure()

        aging_indicators.add_trace(go.Indicator(

            mode = 'number+delta',
            value = cm_aging_90amt,
            delta = {'position': "bottom", 'reference': pm_aging_90amt, 'relative': True, 'valueformat': '.1%', 'increasing': {'color': "red"}, 'decreasing': {'color': "green"}},
            number = {'prefix': '$', 'suffix': 'k'},
            title = {'text': 'Month Over Month 90+'},
            domain = {'x': [0, 0.5], 'y': [0,0.5]}
            ))

        aging_indicators.add_trace(go.Indicator(

            mode = 'number+delta',
            value = cm_aging_90amt,
            delta = {'position': "bottom", 'reference': pm_aging_90amt, 'relative': True, 'valueformat': '.1%', 'increasing': {'color': "red"}, 'decreasing': {'color': "green"}},
            number = {'prefix': '$', 'suffix': 'k'},
            title = {'text':'Week Over Week 90+'},
            domain = {'x': [0.5, 1], 'y': [0.5,1]}
            ))

        # aging_indicators.update_layout(
        #     height = 300,
        #     width = 50
        #     )

        aging_graph = make_subplots(specs=[[{"secondary_y": True}]])
        aging_graph.add_trace(
            go.Scatter(
                x = df_aging_monthly_consolidated_test['Month'],
                y = df_aging_monthly_consolidated_test['90+ % By Consolidated BU'],
                name = '90+ % By BU'),
            secondary_y = True)
        aging_graph.add_trace(
            go.Bar(
                x = df_aging_monthly_consolidated_test['Month'],
                y = df_aging_monthly_consolidated_test['Current'],
                name = 'Current'),
            secondary_y = False)

        aging_graph.add_bar(
            x = df_aging_monthly_consolidated_test['Month'],
            y = df_aging_monthly_consolidated_test['1-30'],
            name = '1-30',
            secondary_y = False)

        aging_graph.add_bar(
            x = df_aging_monthly_consolidated_test['Month'],
            y = df_aging_monthly_consolidated_test['31-60'],
            name = '31-60',
            secondary_y = False)

        aging_graph.add_bar(
            x = df_aging_monthly_consolidated_test['Month'],
            y = df_aging_monthly_consolidated_test['61-90'],
            name = '61-90',
            secondary_y = False)

        aging_graph.add_bar(
            x = df_aging_monthly_consolidated_test['Month'],
            y = df_aging_monthly_consolidated_test['90+'],
            name = '90+',
            secondary_y = False)
        aging_graph.update_layout(
            barmode = 'stack',
            title = ar_aging_title1 + '<br>' + '<span style="font-size: 12px;">By Month</span>' + '<br>' + str(ar_aging_title2),

            )


        df_top5_test = df_top5_test.head(5)

        top5_graph = go.Figure(

            data = [go.Table(
                header = dict(values = list(df_top5_test.columns),
                    fill_color = 'paleturquoise',
                    align = 'left'),
                cells = dict(values=[df_top5_test.columns],
                    fill_color = 'lavender',
                    align = 'left'))
            ])

    else:


        try:
            cm_aging_90amt = df_aging_monthly_region.loc[(df_aging_monthly_region['Consolidated BU'] == option_slctd) & (df_aging_monthly_region['Month'] == cm_aging_month) & (df_aging_monthly_region['Location'] == region_slctd)]
            cm_aging_90amt = cm_aging_90amt.iloc[-1]['90+']

            pm_aging_90amt = df_aging_monthly_region.loc[(df_aging_monthly_region['Consolidated BU'] == option_slctd) & (df_aging_monthly_region['Month'] == pm_aging_month) & (df_aging_monthly_region['Location'] == region_slctd)]
            pm_aging_90amt = pm_aging_90amt.iloc[-1]['90+']
        
        except BaseException:
            cm_aging_90amt = 0
            pm_aging_90amt = 0

        try:
            ar_aging_title3 = ('<span style="font-size: 21px;">Percentage of 90+ Over Total AR = ' + str(round(df_aging_monthly_region_test.iloc[-1]['90+ % By Region Group'].sum()*100,1)) + '%' +'</span>')

        except BaseException:
            ar_aging_title3 = 'No Data Present'

        aging_indicators = go.Figure()

        aging_indicators.add_trace(go.Indicator(

            mode = 'number+delta',
            value = cm_aging_90amt,
            delta = {'position': "bottom", 'reference': pm_aging_90amt, 'relative': True, 'valueformat': '.1%', 'increasing': {'color': "red"}, 'decreasing': {'color': "green"}},
            number = {'prefix': '$', 'suffix': 'k'},
            title = {'text': 'Month Over Month 90+'},
            domain = {'x': [0,0.5], 'y': [0,0]}
            ))

        aging_indicators.add_trace(go.Indicator(

            mode = 'number+delta',
            value = cm_aging_90amt,
            delta = {'position': "bottom", 'reference': pm_aging_90amt, 'relative': True, 'valueformat': '.1%', 'increasing': {'color': "red"}, 'decreasing': {'color': "green"}},
            number = {'prefix': '$', 'suffix': 'k'},
            title = {'text':'Week Over Week 90+'},
            domain = {'x': [0.5,1], 'y': [0,0]}
            ))

        aging_indicators.update_layout(
            height = 300
            )
        # aging_indicators.update_yaxes(
        #     automargin = True
        #     )

        aging_graph = make_subplots(specs=[[{"secondary_y": True}]])
        aging_graph.add_trace(
            go.Scatter(
                x = df_aging_monthly_region_test['Month'],
                y = df_aging_monthly_region_test['90+ % By Region Group'],
                name = '90+ % By BU'),
            secondary_y = True)
        aging_graph.add_trace(
            go.Bar(
                x = df_aging_monthly_region_test['Month'],
                y = df_aging_monthly_region_test['Current'],
                name = 'Current'),
            secondary_y = False)

        aging_graph.add_bar(
            x = df_aging_monthly_region_test['Month'],
            y = df_aging_monthly_region_test['1-30'],
            name = '1-30',
            secondary_y = False)

        aging_graph.add_bar(
            x = df_aging_monthly_region_test['Month'],
            y = df_aging_monthly_region_test['31-60'],
            name = '31-60',
            secondary_y = False)

        aging_graph.add_bar(
            x = df_aging_monthly_region_test['Month'],
            y = df_aging_monthly_region_test['61-90'],
            name = '61-90',
            secondary_y = False)

        aging_graph.add_bar(
            x = df_aging_monthly_region_test['Month'],
            y = df_aging_monthly_region_test['90+'],
            name = '90+',
            secondary_y = False)
        aging_graph.update_layout(barmode = 'stack',
                                  title = ar_aging_title1 + '<br>' + '<span style="font-size: 12px;">By Month</span>' + '<br>' + str(ar_aging_title3),
            )

        df_top5_test = df_top5_test.head(5)
        top5_graph = go.Figure(

            data = [go.Table(
                header = dict(values = list(df_top5_test.columns),
                    fill_color = 'paleturquoise',
                    align = 'left'),
                cells = dict(values=[df_top5_test.columns],
                    fill_color = 'lavender',
                    align = 'left'))
            ])




    return aging_graph, aging_indicators, top5_graph




#------------------------------------------------------------------------------------------------

# Run app
if __name__ == '__main__':
    app.run_server(debug = True,port = 8056)

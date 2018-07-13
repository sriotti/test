########## Import Packages ##########
import dash
import dash_core_components as dcc
import dash_html_components as html
import dash_table_experiments as dt
from dash.dependencies import Input, Output

import numpy as np
import pandas as pd
import datetime
from pyomo.environ import *   # asterisk allows us to import entire pyomo environment to build models
from pyomo.opt import SolverFactory    ##SolverFactory is interface to Solvers
from scipy.stats import norm
from xlsxwriter.utility import xl_rowcol_to_cell
import datetime
import base64
import io

########## Algorithms ##########

def scoping_analysis(demand_df,po_df,input_scope):
    demand_scope = []
    demand_timeframe = demand_df['Transaction Date']> demand_df['Transaction Date'].max() - datetime.timedelta(days=input_scope)
    demand_df = demand_df[demand_timeframe].groupby('Part Number').sum().astype(int)
    demand_df = demand_df[demand_df['Transaction Quantity']>0].sort_values('Transaction Quantity',ascending=False)
    demand_df.reset_index(inplace=True)
    demand_scope = demand_df['Part Number'].values

    po_scope = []
    po_timeframe = po_df['Date Order Placed']> po_df['Date Order Placed'].max() - datetime.timedelta(days=input_scope)
    po_df = po_df[po_timeframe].groupby('Part Number').sum().astype(int)
    po_df = po_df[po_df['Total Spend']>0].sort_values('Total Spend',ascending=False)
    po_df.reset_index(inplace=True)
    po_scope = po_df['Part Number'].values

    inscope_items = list(set(demand_scope).intersection(po_scope))

    model_input_df = pd.DataFrame(data = inscope_items, columns=['Part Number'])

    return model_input_df

def demand_analysis(demand_df, input_scope, model_input_df):
    demand_pivot=demand_df.pivot_table(values=['Transaction Quantity'],
                                   index='Part Number',
                                   columns='Transaction Date',
                                   aggfunc="sum")

    demand_daily_daterange=pd.date_range(start=demand_df['Transaction Date'].max() - datetime.timedelta(days=input_scope - 1),end=demand_df['Transaction Date'].max())

    demand_daily_df=pd.DataFrame(demand_pivot['Transaction Quantity'], model_input_df['Part Number'],demand_daily_daterange.date).fillna(0).astype(int)

    demand_annual = demand_daily_df.sum(axis=1).to_dict()
    model_input_df['Total Demand'] = model_input_df['Part Number'].map(demand_annual)

    demand_daily_average = demand_daily_df.mean(axis=1).to_dict()
    model_input_df['Average Daily Demand'] = model_input_df['Part Number'].map(demand_daily_average)

    demand_daily_variation = demand_daily_df[demand_daily_df >=0].std(axis=1).to_dict()
    model_input_df['Demand Variation'] = model_input_df['Part Number'].map(demand_daily_variation)

    demand_daily_max = demand_daily_df.max(axis=1).to_dict()
    model_input_df['Max Daily Demand'] = model_input_df['Part Number'].map(demand_daily_max)

    return model_input_df

def leadtime_analysis(po_df,input_leadtime_max,model_input_df):
    po_df['Lead Time']=(po_df['Date Order Received']-po_df['Date Order Placed']).dt.days.fillna(0).astype(int)
    po_df.sort_values('Date Order Placed',ascending=False,inplace=True)
    leadtime_criteria = (po_df['Lead Time'] > 0) & (po_df['Lead Time'] < input_leadtime_max)
    leadtime_average = po_df[leadtime_criteria].groupby('Part Number')['Lead Time'].mean().fillna(0).astype(float).to_dict()
    leadtime_variation = po_df[leadtime_criteria].groupby('Part Number')['Lead Time'].std().fillna(0).astype(float).to_dict()
    model_input_df['Average Lead Time'] = model_input_df['Part Number'].map(leadtime_average)
    model_input_df['Lead Time Variation'] = model_input_df['Part Number'].map(leadtime_variation)
    model_input_df.sort_values('Total Demand',ascending=False,inplace=True)

    return model_input_df

def remaining_inputs(demand_df, po_df, snapshot_df, other_df, model_input_df):

    demand_df.sort_values('Transaction Date',ascending=False,inplace=True)
    material_description= demand_df[['Part Number', 'Material Description']].drop_duplicates().set_index('Part Number')['Material Description'].to_dict()
    model_input_df.insert(loc=1, column='Material Description', value=model_input_df['Part Number'].map(material_description))

    po_df.sort_values('Date Order Placed',ascending=False,inplace=True)
    unit_cost = po_df[['Part Number','Unit Cost']].drop_duplicates().set_index('Part Number')['Unit Cost'].to_dict()
    model_input_df.insert(loc=2, column='Unit Cost', value=model_input_df['Part Number'].map(unit_cost))

    DSL={'High': '0.99', 'Medium': '0.98', 'Low': '0.95'}
    criticality = other_df.set_index('Part Number')['Criticality Level'].to_dict()
    criticality = {k: DSL[v] for k, v in criticality.items()}    #dictionary comprehension
    model_input_df.insert(loc=9, column='DSL', value=model_input_df['Part Number'].map(criticality))


    standard_pack_size = other_df.set_index('Part Number')['Standard Pack Size'].to_dict()
    model_input_df.insert(loc=10, column='Standard Pack Size', value=model_input_df['Part Number'].map(standard_pack_size))

    model_input_df['Max Allowed Q'] = model_input_df['Total Demand']
    model_input_df['Inv Review Period Days']= 1
    model_input_df['Minimum Order Frequency Days']= 1
    model_input_df['Shelf Life Days']= 1000000

    snapshot_df['Average'] = snapshot_df.mean(axis=1)
    baseline_units_avg = snapshot_df[['Part Number','Average']].set_index('Part Number')['Average'].to_dict()
    model_input_df['Baseline OH Inv. Units'] = model_input_df['Part Number'].map(baseline_units_avg)
    model_input_df['Baseline OH Inv. Cost'] = model_input_df['Baseline OH Inv. Units'] * model_input_df['Unit Cost']

    model_input_df=model_input_df[(model_input_df['Average Lead Time'] != 0) & (model_input_df['Unit Cost'] != 0)]

    return model_input_df

def model_results(model_input_df, ordering_capacity):
    ######### PYOMO ##########
    model = ConcreteModel()  # Option is concrete vs abstract model. Concrete model you use data now to model v. abstract which you create model first to use on data later on.

    #### Sets

    i = model_input_df['Part Number'].unique()

    #### Parameters

    dailydemand = model_input_df.set_index('Part Number')['Average Daily Demand'].to_dict()
    min_order_quantity = model_input_df.set_index('Part Number')['Standard Pack Size'].to_dict()
    max_order_quantity = model_input_df.set_index('Part Number')['Total Demand'].to_dict()
    min_order_frequency = model_input_df.set_index('Part Number')['Minimum Order Frequency Days'].to_dict()
    shelflife = model_input_df.set_index('Part Number')['Shelf Life Days'].to_dict()
    unitcost = model_input_df.set_index('Part Number')['Unit Cost'].to_dict()

    #### Variables

    model.orders_per_day=Var(i,within = NonNegativeReals)

    #### Constraints

    # def c1_rule(model, i):
    #     return model.orders_per_day[i] <= dailydemand[i]/min_order_quantity[i]
    # model.c1 = Constraint(i,rule = c1_rule)

    def c2_rule(model, i):
        return model.orders_per_day[i] >= dailydemand[i]/max_order_quantity[i]
    model.c2 = Constraint(i,rule = c2_rule)

    def c3_rule(model, i):
        return model.orders_per_day[i] >= 1/shelflife[i]
    model.c3 = Constraint(i,rule = c3_rule)

    def c4_rule(model, i):
        return model.orders_per_day[i] <= 1/min_order_frequency[i]
    model.c4 = Constraint(i,rule = c4_rule)

    def c5_rule(model):
        return sum(model.orders_per_day[x] for x in i) <= ordering_capacity
    model.c5 = Constraint(rule = c5_rule)


    #### Objective Function

    def objective_rule(model):
        return sum(.5*(1/model.orders_per_day[i])*dailydemand[i]*unitcost[i] for i in i)
    model.objective = Objective(rule=objective_rule, sense=minimize, doc='Define objective function')

    def pyomo_postprocess(options=None, instance=None, results=None):
        model.orders_per_day.display()

    #### Solver

    if __name__ == '__main__':
        opt = SolverFactory("ipopt")
        results = opt.solve(model)
        results.write()

    ######## Post Model Run Calc's ########
    dec_vars = []
    for var in model.component_data_objects(Var):
        dec_vars.append(var.parent_component())
    dec_vars = list(set(dec_vars))

    dc = {i : x[i].value
        for x in dec_vars for i in getattr(x, '_index')}

    ##### Append Decision Variable to Results Table
    model_results_df = model_input_df.copy()
    model_results_df.insert(loc=2, column='Model Orders per Day', value=model_results_df['Part Number'].map(dc))

    ##### Post Model Run Calculations
    #Model Order Frequency
    model_results_df['Model Order Frequency'] = 1/model_results_df['Model Orders per Day']
    #Model Recommended Order Quantity
    model_results_df['Model Order Quantity'] = np.maximum(1,round(model_results_df['Model Order Frequency']*model_results_df['Average Daily Demand']/model_results_df['Standard Pack Size'])) * model_results_df['Standard Pack Size']

    #Model Safety Stock
    model_results_df["Model Safety Stock"] = norm.ppf(model_results_df['DSL']) * np.sqrt(
        ((model_results_df['Average Daily Demand']**2) * (model_results_df['Lead Time Variation']**2)) +
        ((model_results_df['Average Lead Time'] + model_results_df['Inv Review Period Days']) * model_results_df['Demand Variation']**2))


        #Model Re-order Point
    model_results_df['Model Re-Order Point'] = round(model_results_df['Model Safety Stock'] + model_results_df['Average Daily Demand'] * (model_results_df['Average Lead Time'] + model_results_df['Inv Review Period Days']))
    #Model OH Inv
    model_results_df['Model OH Inv. Units'] = round(model_results_df['Model Safety Stock'] + model_results_df['Model Order Quantity']/2)
    #Model OH Inv Cost
    model_results_df['Model OH Inv. Cost'] = model_results_df['Model OH Inv. Units'] * model_results_df['Unit Cost']


    model_results_df['Inv. Units Reduction'] = model_results_df['Baseline OH Inv. Units'] - model_results_df['Model OH Inv. Units']
    model_results_df['Inv. Cost Reduction'] = model_results_df['Baseline OH Inv. Cost'] - model_results_df['Model OH Inv. Cost']

    return model_results_df

######## Temporary user inputs ########
input_scope = 365
input_leadtime_max = 140

########## app layout ########

# Setup the app
app = dash.Dash('SPOdashboard')
server = app.server
# wsgi_app = app.wsgi_app

# Boostrap CSS.
app.css.append_css({'external_url': 'https://cdn.rawgit.com/plotly/dash-app-stylesheets/2d266c578d2a6e8850ebce48fdb52759b2aef506/stylesheet-oil-and-gas.css'})  # noqa: E501

#Do i need this?
app.scripts.config.serve_locally = True
html.Div([

])
#Setup the layout of the app
app.layout = html.Div([
    html.Div(
        [
            html.H1(
                'Spare Parts Optimization Model',
                style={'font-family': 'Helvetica',
                       "margin-top": "25",
                       "margin-bottom": "0"},
                className='nine columns',
                ),
            html.Img(
                src="https://2ndvote.com/wp-content/uploads/2016/03/johnson-johnson-logo-1320x364.jpg",
                # src="https://www.cagecode.info/files/cage/300/70628.jpg",
                className='three columns',
                style={
                    'height': '10%',
                    'width': '20%',
                    'float': 'right',
                    'position': 'relative',
                    'padding-top': 10,
                    'padding-right': 0},
                ),
        ],
        className='row'
        ),
    html.P(),
    dcc.Tabs(id="tabs",children=[
        dcc.Tab(label='Summary', children=[
            html.Div(
                [
                    html.P(),
                    html.P(
                        'An Advanced Analytics Model for right-sizing inventory while satisfying desired service level and business constraints.',
                        style={'font-family': 'Helvetica',
                               "font-size": "120%",
                               "width": "80%"},
                        className='eight columns',
                        )

                ]
                )
            ]
            ),
        dcc.Tab(label='Data Upload', children=[
            html.Div(
                [
                    html.Div(
                        [
                            html.P(
                                'Upload Template:',
                                style={
                                    'font-family': 'Helvetica',
                                    "font-size": "100%",
                                    "width": "100%",
                                    'font-weight': 'bold'}
                                ),
                            dcc.Upload(
                                id='upload-data',
                                children=html.Div([
                                    'First Populate Input Fields. Then Drag and Drop or ',
                                    html.A('Select Files')
                                    ]),
                                style={
                                    # 'width': '100%',
                                    'height': '60px',
                                    'lineHeight': '60px',
                                    'borderWidth': '1px',
                                    'borderStyle': 'dashed',
                                    'borderRadius': '5px',
                                    'textAlign': 'center',
                                    'margin-right': '1px',
                                    'margin-top': '1px',
                                    'margin-bottom': '10px'},
                                multiple=True
                                ),
                        ],
                        className='six columns',
                        style={'margin-top':'10'}
                        ),
                    html.Div(
                        [
                            html.P('Enter Data Timeframe (Days):',
                                style={
                                    'font-family': 'Helvetica',
                                    "font-size": "100%",
                                    "width": "100%",
                                    'font-weight': 'bold'
                                },
                            ),
                            dcc.Input(
                                id='input_scope',
                                placeholder='Enter Data Timeframe (Days):',
                                type='number',
                                value='365',
                                style={
                                    'width': '100%',
                                    'height': '60px',
                                    'lineHeight': '60px',
                                    'borderWidth': '1px',
                                    'borderRadius': '5px',
                                    'textAlign': 'center',
                                    'margin': '1px',
                                    'margin-bottom': '10px'
                                },
                            )
                        ],
                        className='two columns',
                        style={'margin-top':'10'}
                        ),
                    html.Div(
                        [
                            html.P('Enter Lead Time Cuttoff (Days):',
                                style={
                                    'font-family': 'Helvetica',
                                    "font-size": "100%",
                                    "width": "100%",
                                    'font-weight': 'bold'
                                },
                            ),
                            dcc.Input(
                                id='input_leadtime_max',
                                placeholder='Enter Lead Time Cuttoff:',
                                type='number',
                                value='150',
                                style={
                                    'width': '100%',
                                    'height': '60px',
                                    'lineHeight': '60px',
                                    'borderWidth': '1px',
                                    'borderRadius': '5px',
                                    'textAlign': 'center',
                                    'margin': '1px',
                                    'margin-bottom': '10px'
                                },
                            )
                        ],
                        className='two columns',
                        style={'margin-top':'10'}
                        ),
                    html.Div(
                        [
                            html.P('Enter Daily Order Capacity (Items):',
                                style={
                                    'font-family': 'Helvetica',
                                    "font-size": "100%",
                                    "width": "100%",
                                    'font-weight': 'bold'
                                },
                            ),
                            dcc.Input(
                                id='input_scope',
                                placeholder='Enter Max Daily Order Capacity (Items):',
                                type='number',
                                value='50',
                                style={
                                    'width': '100%',
                                    'height': '60px',
                                    'lineHeight': '60px',
                                    'borderWidth': '1px',
                                    'borderRadius': '5px',
                                    'textAlign': 'center',
                                    'margin': '1px',
                                    'margin-bottom': '10px'
                                },
                            )
                        ],
                        className='two columns',
                        style={'margin-top':'10'}
                        ),
                    html.Div(
                        [
                            dcc.Dropdown(
                                options=[
                                    {'label': 'Historical Demand Data', 'value': 'demand_df'},
                                    {'label': 'Historical PO Data', 'value': 'po_df'},
                                    {'label': 'Snapshot Summary', 'value': 'snapshot_df'},
                                    {'label': 'Criticality, Standard Pack Size, & Baseline Parameters', 'value': 'other_df'}
                                    ],
                                )
                        ],
                        className='twelve columns',
                        style={'margin-top':'10'}
                        ),
                    html.Div(
                        [
                            html.Div(
                                id='input datatable'
                                ),
                            html.Div(dt.DataTable(rows=[{}]), style={'display': 'none'})
                        ],
                        className='twelve columns',
                        style={'margin-top':'10'}
                        ),
                ],
                className='row'
                )
            ]
            ),
        dcc.Tab(label='Model Inputs', children=[
            html.Div(
                [
                    html.Div(
                        [
                            html.Button('Click to Run Data Analysis and Prepare Optimization Model Inputs',
                                id='button',
                                style={
                                    'width': '100%',
                                    'height': '60px',
                                    'lineHeight': '60px',
                                    'borderWidth': '1px',
                                    'borderRadius': '5px',
                                    'textAlign': 'center',
                                    'margin-left': '1px',
                                    'margin-top': '24px',
                                    'margin-bottom': '10px',
                                    'background-color': '#f44336',
                                    'color': 'white'}
                                )
                        ],
                        className='twleve columns',
                        style={'margin-top':'10'}
                        ),
                ]
                )
            ]
            ),
        dcc.Tab(label='Model Results', children=[
            html.Div(
                [
                    html.Div(
                        [
                            html.Button('Click to Run Optimization Model',
                                id='button',
                                style={
                                    'width': '100%',
                                    'height': '60px',
                                    'lineHeight': '60px',
                                    'borderWidth': '1px',
                                    'borderRadius': '5px',
                                    'textAlign': 'center',
                                    'margin-left': '1px',
                                    'margin-top': '24px',
                                    'margin-bottom': '10px',
                                    'background-color': '#f44336',
                                    'color': 'white'}
                                )
                        ],
                        className='twleve columns',
                        style={'margin-top':'10'}
                        ),
                ]
                )
            ]
            ),
        ],
        # vertical='vertical'
        )
])


######## app interactivity ########

def parse_contents(contents, filename, date,input_scope, input_leadtime_max):
    content_type, content_string = contents.split(',')

    decoded = base64.b64decode(content_string)

    demand_df=pd.read_excel(io.BytesIO(decoded),'Historical Demand',encoding = "ISO-8859-1",parse_dates=['Transaction Date'],converters={'Part Number':str})
    po_df = pd.read_excel(io.BytesIO(decoded),'Historical PO',encoding = "ISO-8859-1",parse_dates=['Date Order Placed', 'Date Order Received'],converters={'Part Number':str, 'Unit Cost':float,'Total Spend':float})
    snapshot_df=pd.read_excel(io.BytesIO(decoded),'Snapshot Summary',converters={'Part Number':str},)
    other_df=pd.read_excel(io.BytesIO(decoded),'Other Data',converters={'Part Number':str})

    model_input_df = scoping_analysis(demand_df,po_df,input_scope)
    model_input_df = demand_analysis(demand_df,input_scope,model_input_df)
    model_input_df = leadtime_analysis(po_df,input_leadtime_max,model_input_df)
    model_input_df = remaining_inputs(demand_df, po_df, snapshot_df, other_df, model_input_df)

    return html.Div(
                [
                    dt.DataTable(
                        rows=model_input_df.to_dict('records'),
                        row_selectable=True,
                        filterable=True,
                        sortable=True,
                        )
                ],
                # className='twelve columns',
                # style={'margin-top':'10'}
                )


@app.callback(Output('input datatable', 'children'),
              [Input('upload-data', 'contents'),
               Input('upload-data', 'filename'),
               Input('upload-data', 'last_modified')])
def update_output(list_of_contents, list_of_names, list_of_dates):
    if list_of_contents is not None:
        children = [
            parse_contents(c, n, d, input_scope, input_leadtime_max) for c, n, d in
            zip(list_of_contents, list_of_names, list_of_dates)]
        return children

######## Create the server ########

app.config['suppress_callback_exceptions']=True

if __name__ == '__main__':
    app.run_server(debug=True)
# def wsgi_app(environ, start_response):
#     status = '200 OK'
#     response_headers = [('Content-type', 'text/plain')]
#     start_response(status, response_headers)
#     response_body = app.layout
#     yield response_body.encode()
#
# if __name__ == '__main__':
#     from wsgiref.simple_server import make_server
#
#     httpd = make_server('localhost', 5555, wsgi_app)
#     httpd.serve_forever()

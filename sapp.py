# Import necessary libraries
import pandas as pd
import numpy as np
from dash import Dash, callback, html, dcc
from dash.dependencies import Input, Output
import plotly.express as px
import requests

# ================= LOAD DATA =================
Orders = pd.read_excel(
    r"C:\Users\Ernest PC\OneDrive\Documents\Superstore.xlsx",
    sheet_name='Orders',
    engine='openpyxl'
)
Returns = pd.read_excel(
    r"C:\Users\Ernest PC\OneDrive\Documents\Superstore.xlsx",
    sheet_name='Returns',
    engine='openpyxl'
)

# Merge Orders and Returns
df = pd.merge(Orders, Returns, how='left', on='Order ID')
pd.set_option('display.max_columns', None)

# Load US States GeoJSON
states_geojson = requests.get(
    "https://raw.githubusercontent.com/PublicaMundi/MappingAPI/master/data/geojson/us-states.json"
).json()

# ================= INITIALIZE APP =================
app = Dash(__name__)

# ================= LAYOUT =================
app.layout = html.Div(children=[
    html.H1(
        "Superstore Sales Analysis",
        style={
            'textAlign': 'center',
            'color': 'darkblue',
            'backgroundColor': 'lightgray',
            'fontFamily': 'Arial-black, sans-serif'
        }
    ),
    # Filters
    html.Div(children=[
        html.Div(children=[
            html.H4("Select a Product Category 👇", style={'marginTop': '20px'}),
            dcc.RadioItems(
                id='product',
                inline=True,
                options=[{'label': 'All Categories', 'value': 'All'}] +
                        [{'label': cat, 'value': cat} for cat in df['Category'].unique()],
                value='All'  # default = All categories
            )
        ], style={'backgroundColor': 'lightgray', 'padding': '5px'}),
        html.Div(children=[
            html.H4('Filter by Region'),
            dcc.Dropdown(
                id='region',
                options=[{'label': 'All Regions', 'value': 'All'}] +
                        [{'label': region, 'value': region} for region in df['Region'].unique()],
                value=['All'],  # default = All regions
                multi=True
            )
        ], style={'backgroundColor': 'lightgray', 'padding': '5px', 'width': '40%'})
    ], style={'display': 'flex', 'gap': '20px'}),
    
    # Graphs row 1
    html.Div(
        children=[
            dcc.Graph(id='bar-graph', figure={}, style={'height': '500px', 'width': '50%'}),
            dcc.Graph(id='bar-graph1', figure={}, style={'height': '500px', 'width': '50%'})
        ],
        style={'display': 'flex', 'justifyContent': 'space-around'}
    ),
    # Graphs row 2
    html.Div(children=[
        dcc.Graph(id='sunburst-graph', figure={}, style={'height': '500px', 'width': '50%'}),
        dcc.Graph(id='map-graph', figure={}, style={'height': '500px', 'width': '50%'})
    ], style={'display': 'flex', 'justifyContent': 'space-around'})
])

# ================= CALLBACK =================
@callback(
    Output('bar-graph', 'figure'),
    Output('bar-graph1', 'figure'),
    Output('sunburst-graph', 'figure'),
    Output('map-graph', 'figure'),
    Input('product', 'value'),
    Input('region', 'value')
)
def update_graph(selected_category, selected_region):
    # --- Apply filters ---
    if selected_category == 'All':
        category_filter = df
    else:
        category_filter = df[df['Category'] == selected_category]

    if 'All' in selected_region:
        filtered_df = category_filter
    else:
        filtered_df = category_filter[category_filter['Region'].isin(selected_region)]

    # --- Monthly Sales Chart ---
    monthly_sales = (
        filtered_df.groupby(filtered_df['Order Date'].dt.to_period("M"))['Sales']
        .sum()
        .reset_index()
    )
    monthly_sales['Order Date'] = monthly_sales['Order Date'].dt.to_timestamp()

    fig = px.line(
        monthly_sales,
        x='Order Date',
        y='Sales',
        title=f"Monthly Sales ({selected_category if selected_category != 'All' else 'All Categories'})",
        labels={'Order Date': 'Month', 'Sales': 'Total Sales'}
    )

    # --- Subcategory Bar Chart ---
    sub_sales = (
        filtered_df.groupby('Sub-Category')['Sales']
        .sum()
        .reset_index()
    )
    fig1 = px.bar(
        sub_sales,
        x='Sub-Category',
        y='Sales',
        title=f"Total Sales by Sub-Category",
        labels={'Sub-Category': 'Sub-category', 'Sales': 'Total Sales'},
        color='Sub-Category'
    )

    # --- Sunburst Chart ---
    sunburst_fig = px.sunburst(
        filtered_df,
        path=['Region', 'State', 'City'],
        values='Sales',
        title="Sales Distribution by Region → State → City"
    )
    sunburst_fig.update_traces(
        textinfo="label+percent entry+value",
        insidetextorientation="radial",
        hovertemplate="<b>%{label}</b><br>Parent: %{parent}<br>Sales: %{value:$,.0f}<br>Share: %{percentEntry:.2%}<extra></extra>"
    )

    # --- US State Map ---
    state_sales = filtered_df.groupby("State", as_index=False)["Sales"].sum()
    map_fig = px.choropleth_mapbox(
        state_sales,
        geojson=states_geojson,
        locations="State",
        featureidkey="properties.name",
        color="Sales",
        color_continuous_scale="Blues",
        mapbox_style="carto-positron",
        zoom=2.3,
        center={"lat": 37.0902, "lon": -95.7129},
        opacity=0.6,
        title="Sales by State"
    )

    return fig, fig1, sunburst_fig, map_fig

# ================= RUN APP =================
if __name__ == "__main__":
    app.run(debug=True)

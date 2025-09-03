# Import necessary libraries
import pandas as pd
from dash import Dash, callback, html, dcc, ctx
from dash.dependencies import Input, Output, State
import plotly.express as px
import json
import requests
from datetime import datetime
from pathlib import Path

path = Path(__file__).parent.parent / 'data' / 'Superstore.xlsx'

# Check if the file exists before loading
if not path.exists():
    raise FileNotFoundError(f"Excel file not found at {path}")

# Load the dataset
Orders = pd.read_excel(str(path), sheet_name='Orders', engine='openpyxl')
Returns = pd.read_excel(str(path), sheet_name='Returns', engine='openpyxl')

# Merge Orders and Returns
df = pd.merge(Orders, Returns, how='left', on='Order ID')
pd.set_option('display.max_columns', None)

# Load US States GeoJSON from Plotly
us_states_url = "https://raw.githubusercontent.com/plotly/datasets/master/geojson-counties-fips.json"
states_geojson = requests.get("https://raw.githubusercontent.com/PublicaMundi/MappingAPI/master/data/geojson/us-states.json").json()

# Initialize Dash app
app = Dash(__name__)

# Layout
app.layout = html.Div(children=[
    html.H1(
        "Interactive Superstore Sales Analysis",
        style={
            'textAlign': 'center',
            'color': 'darkblue',
            'backgroundColor': 'lightgray',
            'fontFamily': 'Arial-black, sans-serif',
            'padding': '20px',
            'margin': '0 0 20px 0'
        }
    ),
    
    # Filters section
    html.Div(children=[
        html.Div(children=[
            html.H4("Select Product Category 👇", style={'marginTop': '0px', 'color': 'darkblue'}),
            dcc.RadioItems(
                id='product',
                inline=True,
                options=[{'label': 'All Categories', 'value': 'All'}] +
                        [{'label': cat, 'value': cat} for cat in df['Category'].unique()],
                value='All',
                style={'fontSize': '14px'}
            )
        ], style={'backgroundColor': 'lightgray', 'padding': '15px', 'margin': '5px', 'borderRadius': '5px', 'width': '48%'}),
        
        html.Div(children=[
            html.H4('Filter by Region 🌎', style={'marginTop': '0px', 'color': 'darkblue'}),
            dcc.Dropdown(
                id='region',
                options=[{'label': 'All Regions', 'value': 'All'}] +
                        [{'label': region, 'value': region} for region in df['Region'].unique()],
                value=['All'],
                multi=True,
                placeholder="Select regions..."
            )
        ], style={'backgroundColor': 'lightgray', 'padding': '15px', 'margin': '5px', 'borderRadius': '5px', 'width': '48%'})
    ], style={'display': 'flex', 'gap': '10px', 'marginBottom': '20px'}),
    
    # Store components for cross-filtering
    dcc.Store(id='selected-state'),
    dcc.Store(id='selected-subcategory'),
    
    # Clear filters button
    html.Div([
        html.Button('Clear All Filters', id='clear-filters-btn', 
                   style={'backgroundColor': '#ff6b6b', 'color': 'white', 'border': 'none', 
                         'padding': '10px 20px', 'borderRadius': '5px', 'cursor': 'pointer',
                         'fontSize': '14px', 'marginBottom': '20px'})
    ]),
    
    # First row of charts
    html.Div(
        children=[
            html.Div([
                dcc.Graph(id='bar-graph', figure={}, 
                         style={'height': '500px'},
                         config={'displayModeBar': True}),
                html.P("📊 Click on data points to filter other charts", 
                      style={'textAlign': 'center', 'color': 'gray', 'fontSize': '12px'})
            ], style={'width': '50%', 'padding': '5px'}),
            
            html.Div([
                dcc.Graph(id='bar-graph1', figure={}, 
                         style={'height': '500px'},
                         config={'displayModeBar': True}),
                html.P("📈 Click on subcategories to filter by selection", 
                      style={'textAlign': 'center', 'color': 'gray', 'fontSize': '12px'})
            ], style={'width': '50%', 'padding': '5px'})
        ],
        style={'display': 'flex', 'justifyContent': 'space-around'}
    ),
    
    # Second row of charts
    html.Div(children=[
        html.Div([
            dcc.Graph(id='sunburst-graph', figure={}, 
                     style={'height': '500px'},
                     config={'displayModeBar': True}),
            html.P("🌞 Explore hierarchy by clicking on segments", 
                  style={'textAlign': 'center', 'color': 'gray', 'fontSize': '12px'})
        ], style={'width': '50%', 'padding': '5px'}),
        
        html.Div([
            dcc.Graph(id='map-graph', figure={}, 
                     style={'height': '500px'},
                     config={'displayModeBar': True}),
            html.P("🗺️ Click on states to filter data by location", 
                  style={'textAlign': 'center', 'color': 'gray', 'fontSize': '12px'})
        ], style={'width': '50%', 'padding': '5px'})
    ], style={'display': 'flex', 'justifyContent': 'space-around'}),
    
    # Selected filters display
    html.Div(id='current-filters', style={'marginTop': '20px', 'padding': '15px', 
                                         'backgroundColor': '#f0f0f0', 'borderRadius': '5px'})
])

# Callback for clearing filters
@callback(
    Output('product', 'value'),
    Output('region', 'value'),
    Output('selected-state', 'data'),
    Output('selected-subcategory', 'data'),
    Input('clear-filters-btn', 'n_clicks'),
    prevent_initial_call=True
)
def clear_filters(n_clicks):
    if n_clicks:
        return 'All', ['All'], None, None
    return 'All', ['All'], None, None

# Callback for map clicks
@callback(
    Output('selected-state', 'data', allow_duplicate=True),
    Input('map-graph', 'clickData'),
    prevent_initial_call=True
)
def update_selected_state(clickData):
    if clickData is None:
        return None
    return clickData['points'][0]['location']

# Callback for subcategory bar clicks
@callback(
    Output('selected-subcategory', 'data', allow_duplicate=True),
    Input('bar-graph1', 'clickData'),
    prevent_initial_call=True
)
def update_selected_subcategory(clickData):
    if clickData is None:
        return None
    return clickData['points'][0]['x']

# Main callback to update all graphs
@callback(
    Output('bar-graph', 'figure'),
    Output('bar-graph1', 'figure'),
    Output('sunburst-graph', 'figure'),
    Output('map-graph', 'figure'),
    Output('current-filters', 'children'),
    Input('product', 'value'),
    Input('region', 'value'),
    Input('selected-state', 'data'),
    Input('selected-subcategory', 'data')
)
def update_graphs(selected_category, selected_regions, selected_state, selected_subcategory):
    # Start with full dataset
    filtered_df = df.copy()
    filter_info = []
    
    # Apply category filter
    if selected_category != 'All':
        filtered_df = filtered_df[filtered_df['Category'] == selected_category]
        filter_info.append(f"Category: {selected_category}")
    
    # Apply region filter
    if selected_regions and 'All' not in selected_regions:
        filtered_df = filtered_df[filtered_df['Region'].isin(selected_regions)]
        filter_info.append(f"Regions: {', '.join(selected_regions)}")
    
    # Apply state filter (from map clicks)
    if selected_state:
        filtered_df = filtered_df[filtered_df['State'] == selected_state]
        filter_info.append(f"State: {selected_state}")
    
    # Apply subcategory filter (from bar chart clicks)
    if selected_subcategory:
        filtered_df = filtered_df[filtered_df['Sub-Category'] == selected_subcategory]
        filter_info.append(f"Sub-Category: {selected_subcategory}")
    
    # ================= MONTHLY SALES CHART =================
    if len(filtered_df) > 0:
        monthly_sales = (
            filtered_df.groupby([filtered_df['Order Date'].dt.to_period("M"), 'Category'])['Sales']
            .sum()
            .reset_index()
        )
        monthly_sales['Order Date'] = monthly_sales['Order Date'].dt.to_timestamp()
        
        if selected_category == 'All':
            fig = px.line(
                monthly_sales,
                x='Order Date',
                y='Sales',
                color='Category',
                title="📈 Monthly Sales Trend by Category",
                labels={'Order Date': 'Month', 'Sales': 'Total Sales ($)'}
            )
        else:
            fig = px.bar(
                monthly_sales,
                x='Order Date',
                y='Sales',
                color='Category',
                title=f"📊 Monthly Sales for {selected_category}",
                labels={'Order Date': 'Month', 'Sales': 'Total Sales ($)'}
            )
    else:
        fig = px.bar(title="No data available for current filters")
    
    fig.update_layout(hovermode='x unified', showlegend=True)
    
    # ================= SUBCATEGORY BAR CHART =================
    if len(filtered_df) > 0:
        subcategory_sales = (
            filtered_df.groupby(['Sub-Category', 'Category'])['Sales']
            .sum()
            .reset_index()
            .sort_values('Sales', ascending=True)
        )
        
        fig1 = px.bar(
            subcategory_sales,
            x='Sales',
            y='Sub-Category',
            color='Category',
            orientation='h',
            title="💼 Sales by Sub-Category",
            labels={'Sub-Category': 'Sub-Category', 'Sales': 'Total Sales ($)'}
        )
        
        # Add click instructions
        fig1.update_layout(
            hovermode='y unified',
            height=500
        )
    else:
        fig1 = px.bar(title="No data available for current filters")
    
    # ================= SUNBURST CHART =================
    if len(filtered_df) > 0:
        sunburst_fig = px.sunburst(
            filtered_df,
            path=['Region', 'State', 'Category', 'Sub-Category'],
            values='Sales',
            title="🌞 Sales Hierarchy: Region → State → Category → Sub-Category"
        )
        
        sunburst_fig.update_traces(
            textinfo="label+percent entry",
            insidetextorientation="radial",
            hovertemplate=(
                "<b>%{label}</b><br>"
                "Parent: %{parent}<br>"
                "Sales: $%{value:,.0f}<br>"
                "Share: %{percentEntry:.1%}<extra></extra>"
            )
        )
    else:
        sunburst_fig = px.sunburst(title="No data available for current filters")
    
    # ================= US STATE MAP =================
    if len(filtered_df) > 0:
        state_sales = filtered_df.groupby("State", as_index=False)["Sales"].sum()
        
        map_fig = px.choropleth_mapbox(
            state_sales,
            geojson=states_geojson,
            locations="State",
            featureidkey="properties.name",
            color="Sales",
            color_continuous_scale="Viridis",
            mapbox_style="carto-positron",
            zoom=2.3,
            center={"lat": 37.0902, "lon": -95.7129},
            opacity=0.7,
            title="🗺️ Sales by State (Click to filter)",
            hover_data={'Sales': ':$,.0f'}
        )
        
        map_fig.update_layout(
            coloraxis_colorbar=dict(title="Sales ($)")
        )
    else:
        map_fig = px.choropleth_mapbox(
            mapbox_style="carto-positron",
            zoom=2.3,
            center={"lat": 37.0902, "lon": -95.7129},
            title="No data available for current filters"
        )
    
    # ================= FILTER DISPLAY =================
    if filter_info:
        filter_display = html.Div([
            html.H4("🔍 Active Filters:", style={'color': 'darkblue', 'margin': '0 0 10px 0'}),
            html.Ul([html.Li(info, style={'color': 'darkgreen'}) for info in filter_info])
        ])
    else:
        filter_display = html.Div([
            html.H4("🔍 No filters applied", style={'color': 'gray', 'margin': '0'}),
            html.P("Click on charts to apply cross-filters", style={'color': 'gray', 'margin': '5px 0 0 0'})
        ])
    
    return fig, fig1, sunburst_fig, map_fig, filter_display

# Run the app
if __name__ == "__main__":
    app.run(debug=True, port=8051)
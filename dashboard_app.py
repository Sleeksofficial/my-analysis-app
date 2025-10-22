import dash
from dash import dcc, html, Input, Output, State, callback, clientside_callback
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import base64
import io

# Import the functions from your other file
from analysis_module import analyze_excel, generate_report
import excel_exporter

# --- NEW: Import forecasting library ---
# try:
#     from statsmodels.tsa.holtwinters import ExponentialSmoothing
# except ImportError:
#     print("WARNING: statsmodels not found. Forecasting will not work.")
ExponentialSmoothing = None # Force to None

# Initialize the Dash app
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP, dbc.icons.BOOTSTRAP])
server = app.server

# --- NEW: Navbar Layout ---
navbar = dbc.Navbar(
    dbc.Container([
        html.A(
            dbc.Row(
                [
                    dbc.Col(html.I(className="bi bi-bar-chart-line-fill", style={'color': '#0d6efd', 'fontSize': 30})),
                    dbc.Col(dbc.NavbarBrand("AI Data Dashboard", className="ms-2")),
                ],
                align="center",
                className="g-0",
            ),
            href="/",
            style={"textDecoration": "none"},
        ),
        dbc.Row([
            dbc.Col(dcc.Dropdown(
                id='theme-selector', 
                placeholder="Select Theme...",
                options=[
                    {'label': 'Default', 'value': 'plotly'},
                    {'label': 'Simple', 'value': 'simple_white'},
                    {'label': 'Dark', 'value': 'plotly_dark'},
                    {'label': 'Seaborn', 'value': 'seaborn'},
                ],
                value='plotly'
            ), className="no-print", style={'width': '200px'}),
            dbc.Col(dbc.Button("Export to Excel", id="btn-export-excel", color="success", n_clicks=0, className="ms-2 no-print")),
            dbc.Col(dbc.Button("Export to PDF", id="btn-print-pdf", color="primary", n_clicks=0, className="ms-2 no-print")),
        ], className="g-0 ms-auto flex-nowrap mt-3 mt-md-0", align="center")
    ]),
    color="light",
    dark=False,
    sticky="top",
    className="no-print"
)

# --- NEW: Attractive Homepage Layout ---
homepage_layout = dbc.Container([
    dbc.Row(
        dbc.Col(
            html.H1("Dynamic Data Dashboard", className="text-center text-light mt-5"),
            width=12
        ),
        className="mb-4"
    ),
    dbc.Row(
        dbc.Col(
            dbc.Card([
                dbc.CardHeader(html.H4("Start Your Analysis", className="text-center")),
                dbc.CardBody([
                    
                    # --- NEW: Error Alert (with correct closing parenthesis) ---
                    dbc.Alert(
                        id="upload-error-alert",
                        color="danger",
                        is_open=False, # Start hidden
                        className="no-print"
                    ), # <-- This closing parenthesis was likely missing
                    
                    dcc.Upload(
                        id='upload-data', # This is the trigger
                        children=html.Div([
                            'Drag and Drop or ',
                            html.A('Select Your Excel File')
                        ]),
                        style={
                            'width': '100%', 'height': '120px', 'lineHeight': '120px',
                            'borderWidth': '2px', 'borderStyle': 'dashed',
                            'borderRadius': '5px', 'textAlign': 'center', 'margin': '10px 0'
                        },
                        multiple=False
                    ),
                    dbc.Alert([
                        html.I(className="bi bi-info-circle-fill me-2"),
                        "Upload an Excel file (.xlsx, .xls) to instantly generate a detailed report and interactive dashboard."
                    ], color="primary", className="mt-3")
                ])
            ]),
            width={'size': 8, 'offset': 2} # Center the card
        )
    )
], fluid=True, className="vh-100 animated-gradient-bg", id="homepage-container") # ID is key
# --- NEW: Main Dashboard Layout ---
main_dashboard_layout = dbc.Container([
    dcc.Download(id="download-excel"),
    
    # --- NEW: Toast Notification Area ---
    html.Div(
        dbc.Toast(
            id="excel-toast",
            header="Generating Report",
            icon="primary",
            duration=4000,
            is_open=False,
            style={"position": "fixed", "top": 66, "right": 10, "width": 350, "zIndex": 9999},
        ),
    ),
    
    # --- Control Row ---
    dbc.Row([
        dbc.Col(dcc.Dropdown(id='sheet-selector-dropdown', placeholder="Select Sheet..."), md=12)
    ], align="center", className="mb-3 no-print"),
    
    html.Hr(className="no-print"),
    
    # --- Dashboard Tabs ---
    dbc.Tabs(id="dashboard-tabs", active_tab="tab-0", children=[
        
        # --- Tab 1: Report ---
        dbc.Tab(label="Detailed Analysis Report", tab_id="tab-0", children=[
            dcc.Loading(dbc.Row(dbc.Col(dcc.Markdown(id='analysis-report', className="mt-4"), width=12)))
        ]),
        
        # --- NEW Tab 2: Data Cleaning ---
        dbc.Tab(label="Data Cleaning", tab_id="tab-1", children=[
            dbc.Row([
                dbc.Col(dbc.Card([
                    dbc.CardHeader("Cleaning Options"),
                    dbc.CardBody([
                        dbc.Switch(label="Remove Duplicate Rows", id="clean-duplicates-switch", value=False),
                        dbc.Switch(label="Drop Rows with Missing Values", id="clean-na-switch", value=False),
                        html.P("Changes here will re-run the analysis and update all other tabs.", className="text-muted small mt-3")
                    ])
                ]), md=4),
                dbc.Col(dbc.Card([
                    dbc.CardHeader("Data Preview (After Cleaning)"),
                    dbc.CardBody(id="data-preview-table")
                ]), md=8)
            ], className="mt-4")
        ]),

        # --- NEW Tab 3: Correlation Matrix ---
        dbc.Tab(label="Correlation Matrix", tab_id="tab-2", children=[
            dcc.Loading(dcc.Graph(id='correlation-heatmap', style={'height': '70vh'}))
        ]),

        # --- NEW Tab 4: Forecasting ---
        # dbc.Tab(label="Forecasting", tab_id="tab-3", children=[
        #     dbc.Card(dbc.CardBody([
        #         dbc.Row([
        #             dbc.Col(dcc.Dropdown(id='forecast-date-col', placeholder="Select Date Column..."), width=4),
        #             dbc.Col(dcc.Dropdown(id='forecast-metric-col', placeholder="Select Metric to Forecast..."), width=4),
        #             dbc.Col(dbc.Button("Run Forecast", id="run-forecast-button", n_clicks=0, color="primary"), width=4),
        #         ]),
        #         dcc.Loading(dbc.Graph(id='forecast-graph', style={'height': '60vh'}))
        #     ]), className="mt-4")
        # ]),
        
        # --- Tab 5: Interactive Dashboard ---
        dbc.Tab(label="Interactive Dashboard", tab_id="tab-4", children=[
            dbc.Row([
                dbc.Col(md=4, className="mb-4", children=[
                    dbc.Card([
                        dbc.CardBody([
                            html.H5("Chart 1", className="card-title no-print"),
                            dcc.Dropdown(id='chart1-type', placeholder="Select Chart Type", options=['Bar Chart', 'Line Chart', 'Scatter Plot', 'Histogram', 'Pie Chart'], className="no-print"),
                            dcc.Dropdown(id='chart1-x', placeholder="Select X-Axis", className="no-print"),
                            dcc.Dropdown(id='chart1-y', placeholder="Select Y-Axis (Optional)", className="no-print"),
                            dcc.Dropdown(id='chart1-color', placeholder="Select Color/Group (Optional)", className="no-print"),
                            dcc.Loading(dcc.Graph(id='graph1', style={'height': '450px'}))
                        ])
                    ])
                ]),
                dbc.Col(md=4, className="mb-4", children=[
                    dbc.Card([
                        dbc.CardBody([
                            html.H5("Chart 2", className="card-title no-print"),
                            dcc.Dropdown(id='chart2-type', placeholder="Select Chart Type", options=['Bar Chart', 'Line Chart', 'Scatter Plot', 'Histogram', 'Pie Chart'], className="no-print"),
                            dcc.Dropdown(id='chart2-x', placeholder="Select X-Axis", className="no-print"),
                            dcc.Dropdown(id='chart2-y', placeholder="Select Y-Axis (Optional)", className="no-print"),
                            dcc.Dropdown(id='chart2-color', placeholder="Select Color/Group (Optional)", className="no-print"),
                            dcc.Loading(dcc.Graph(id='graph2', style={'height': '450px'}))
                        ])
                    ])
                ]),
                dbc.Col(md=4, className="mb-4", children=[
                    dbc.Card([
                        dbc.CardBody([
                            html.H5("Chart 3", className="card-title no-print"),
                            dcc.Dropdown(id='chart3-type', placeholder="Select Chart Type", options=['Bar Chart', 'Line Chart', 'Scatter Plot', 'Histogram', 'Pie Chart'], className="no-print"),
                            dcc.Dropdown(id='chart3-x', placeholder="Select X-Axis", className="no-print"),
                            dcc.Dropdown(id='chart3-y', placeholder="Select Y-Axis (Optional)", className="no-print"),
                            dcc.Dropdown(id='chart3-color', placeholder="Select Color/Group (Optional)", className="no-print"),
                            dcc.Loading(dcc.Graph(id='graph3', style={'height': '450px'}))
                        ])
                    ])
                ]),
            ], className="mt-4"),
            dbc.Row([
                dbc.Col(md=4, className="mb-4", children=[
                    dbc.Card([
                        dbc.CardBody([
                            html.H5("Chart 4", className="card-title no-print"),
                            dcc.Dropdown(id='chart4-type', placeholder="Select Chart Type", options=['Bar Chart', 'Line Chart', 'Scatter Plot', 'Histogram', 'Pie Chart'], className="no-print"),
                            dcc.Dropdown(id='chart4-x', placeholder="Select X-Axis", className="no-print"),
                            dcc.Dropdown(id='chart4-y', placeholder="Select Y-Axis (Optional)", className="no-print"),
                            dcc.Dropdown(id='chart4-color', placeholder="Select Color/Group (Optional)", className="no-print"),
                            dcc.Loading(dcc.Graph(id='graph4', style={'height': '450px'}))
                        ])
                    ])
                ]),
                dbc.Col(md=4, className="mb-4", children=[
                    dbc.Card([
                        dbc.CardBody([
                            html.H5("Chart 5", className="card-title no-print"),
                            dcc.Dropdown(id='chart5-type', placeholder="Select Chart Type", options=['Bar Chart', 'Line Chart', 'Scatter Plot', 'Histogram', 'Pie Chart'], className="no-print"),
                            dcc.Dropdown(id='chart5-x', placeholder="Select X-Axis", className="no-print"),
                            dcc.Dropdown(id='chart5-y', placeholder="Select Y-Axis (Optional)", className="no-print"),
                            dcc.Dropdown(id='chart5-color', placeholder="Select Color/Group (Optional)", className="no-print"),
                            dcc.Loading(dcc.Graph(id='graph5', style={'height': '450px'}))
                        ])
                    ])
                ]),
                dbc.Col(md=4, className="mb-4", children=[
                    dbc.Card([
                        dbc.CardBody([
                            html.H5("Chart 6", className="card-title no-print"),
                            dcc.Dropdown(id='chart6-type', placeholder="Select Chart Type", options=['Bar Chart', 'Line Chart', 'Scatter Plot', 'Histogram', 'Pie Chart'], className="no-print"),
                            dcc.Dropdown(id='chart6-x', placeholder="Select X-Axis", className="no-print"),
                            dcc.Dropdown(id='chart6-y', placeholder="Select Y-Axis (Optional)", className="no-print"),
                            dcc.Dropdown(id='chart6-color', placeholder="Select Color/Group (Optional)", className="no-print"),
                            dcc.Loading(dcc.Graph(id='graph6', style={'height': '450px'}))
                        ])
                    ])
                ]),
            ], className="mt-4")
        ]),
    ]),
], fluid=True, style={'display': 'none'}, id="main-dashboard-wrapper") # Starts hidden


# --- Main App Layout ---
app.layout = html.Div([
    dcc.Store(id='data-is-loaded', data=False),
    dcc.Store(id='stored-data-summary'),
    dcc.Store(id='stored-data-sheet-options'), # This holds the RAW data
    dcc.Store(id='cleaned-data-store'),       # This holds the CLEANED data
    
    navbar, # The Navbar is always visible
    homepage_layout,
    main_dashboard_layout
])


# --- Callbacks ---

def parse_contents(contents):
    """Helper function to parse the uploaded file."""
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    return io.BytesIO(decoded)

# --- NEW: This callback handles switching between the homepage and the dashboard ---
@callback(
    [Output('homepage-container', 'style'),
     Output('main-dashboard-wrapper', 'style'),
     Output('data-is-loaded', 'data'),
     Output('stored-data-summary', 'data'),
     Output('stored-data-sheet-options', 'data'),
     Output('upload-error-alert', 'children'),    # --- ADDED ---
     Output('upload-error-alert', 'is_open')],   # --- ADDED ---
    [Input('upload-data', 'contents')]
)
def handle_file_upload(contents):
    if contents is None:
        # No upload yet, show homepage and hide dashboard
        return {'display': 'block'}, {'display': 'none'}, False, dash.no_update, dash.no_update, None, False

    file_buffer = parse_contents(contents)
    
    # --- FIX: Check for the new error_message ---
    summary, data, error_message = analyze_excel(file_buffer)
    
    if summary is None or error_message:
        # Upload failed, stay on homepage and show the error
        error_text = f"File Upload Failed: {error_message}"
        return {'display': 'block'}, {'display': 'none'}, False, None, None, error_text, True

    # Upload succeeded!
    json_data = {sheet: df.to_json(orient='split') for sheet, df in data.items()}
    
    # Hide homepage, show dashboard, and store the data
    return {'display': 'none'}, {'display': 'block'}, True, summary, json_data, None, False

# --- NEW: Callback to clean the data ---
@callback(
    Output('cleaned-data-store', 'data'),
    [Input('stored-data-sheet-options', 'data'), # Triggered when raw data is loaded
     Input('clean-duplicates-switch', 'value'),  # Triggered by cleaning toggle
     Input('clean-na-switch', 'value')]          # Triggered by cleaning toggle
)
def clean_data(json_data, remove_duplicates, drop_na):
    if not json_data:
        return None
    
    cleaned_data = {}
    for sheet, df_json in json_data.items():
        df = pd.read_json(io.StringIO(df_json), orient='split')
        
        if remove_duplicates:
            df = df.drop_duplicates()
        if drop_na:
            df = df.dropna()
            
        cleaned_data[sheet] = df.to_json(orient='split')
    
    return cleaned_data

# --- MODIFIED: This callback populates all tabs based on the CLEANED data ---
@callback(
    [Output('sheet-selector-dropdown', 'options'),
     Output('sheet-selector-dropdown', 'value'),
     Output('analysis-report', 'children'),
     Output('data-preview-table', 'children'),
     # Chart Dropdowns (Smarter)
     Output('chart1-x', 'options'), Output('chart1-y', 'options'), Output('chart1-color', 'options'),
     Output('chart2-x', 'options'), Output('chart2-y', 'options'), Output('chart2-color', 'options'),
     Output('chart3-x', 'options'), Output('chart3-y', 'options'), Output('chart3-color', 'options'),
     Output('chart4-x', 'options'), Output('chart4-y', 'options'), Output('chart4-color', 'options'),
     Output('chart5-x', 'options'), Output('chart5-y', 'options'), Output('chart5-color', 'options'),
     Output('chart6-x', 'options'), Output('chart6-y', 'options'), Output('chart6-color', 'options')],
     # Forecasting outputs were removed, so this is the complete list
    [Input('cleaned-data-store', 'data')], # <-- ONLY triggered by clean data
    [State('stored-data-summary', 'data')]  # <-- Get summary as State
)
def update_all_tabs_from_cleaned_data(cleaned_json_data, summary):
    if not cleaned_json_data or not summary:
        # 4 outputs + 18 chart outputs = 22 total
        return ([], None, "Please upload a file to begin.", None) + ([[]]*18)

    # Rebuild dataframes from the *cleaned* json
    data_dfs = {sheet: pd.read_json(io.StringIO(cleaned_json_data[sheet]), orient='split') for sheet in cleaned_json_data}
    
    # --- 1. Populate Sheet Selector ---
    sheet_options = [{'label': sheet, 'value': sheet} for sheet in summary.keys()]
    selected_sheet = sheet_options[0]['value']
    
    # --- 2. Generate AI Report ---
    report = generate_report(summary, data_dfs)
    
    # --- 3. Generate Data Preview Table ---
    df_preview = data_dfs[selected_sheet]
    preview_table = dbc.Table.from_dataframe(df_preview.head(5), striped=True, bordered=True, hover=True, responsive=True)

    # --- 4. Populate Dropdowns (SMARTER) ---
    sheet_summary = summary[selected_sheet]
    all_cols = [{'label': col, 'value': col} for col in sheet_summary['Columns']]
    num_cols = [{'label': col, 'value': col} for col in sheet_summary['Numeric_Columns']]
    cat_cols = [{'label': col, 'value': col} for col in sheet_summary['Categorical_Columns']]
    
    # (Bar, Line, Scatter, Hist, Pie, Other)
    dropdown_options = (
        cat_cols, num_cols, cat_cols,  # Chart 1 (Bar)
        all_cols, num_cols, cat_cols,  # Chart 2 (Line)
        num_cols, num_cols, cat_cols,  # Chart 3 (Scatter)
        num_cols, num_cols, cat_cols,  # Chart 4 (Histogram)
        cat_cols, num_cols, cat_cols,  # Chart 5 (Pie)
        all_cols, num_cols, cat_cols   # Chart 6 (Flexible)
    )
    
    # Return 4 main items + 18 dropdown items = 22 total
    return (sheet_options, selected_sheet, report, preview_table) + dropdown_options

# --- NEW: Callback for Correlation Heatmap ---
@callback(
    Output('correlation-heatmap', 'figure'),
    [Input('dashboard-tabs', 'active_tab'),
     Input('sheet-selector-dropdown', 'value'),
     Input('theme-selector', 'value')],
    [State('cleaned-data-store', 'data')]
)
def update_correlation_heatmap(active_tab, selected_sheet, template, cleaned_json_data):
    if active_tab != 'tab-2' or not selected_sheet or not cleaned_json_data:
        return go.Figure()
        
    df = pd.read_json(io.StringIO(cleaned_json_data[selected_sheet]), orient='split')
    num_cols = df.select_dtypes(include='number').columns.tolist()
    
    if len(num_cols) < 2:
        return go.Figure().update_layout(title="Not enough numeric data for correlation", template=template)
    
    corr = df[num_cols].corr()
    fig = px.imshow(corr, text_auto=True, aspect="auto",
                    title=f"Correlation Heatmap for {selected_sheet}",
                    template=template, color_continuous_scale='RdBu_r')
    return fig

# --- NEW: Callback for Forecasting ---
# @callback(
#     Output('forecast-graph', 'figure'),
#     [Input('run-forecast-button', 'n_clicks')],
#     [State('forecast-date-col', 'value'),
#      State('forecast-metric-col', 'value'),
#      State('sheet-selector-dropdown', 'value'),
#      State('theme-selector', 'value'),
#      State('cleaned-data-store', 'data')]
# )
# def run_forecast(n_clicks, date_col, metric_col, selected_sheet, template, cleaned_json_data):
#     if n_clicks == 0 or not all([date_col, metric_col, selected_sheet, cleaned_json_data]):
#         return go.Figure().update_layout(title="Select a date, metric, and run forecast", template=template)
        
#     if ExponentialSmoothing is None:
#         return go.Figure().update_layout(title="Error: 'statsmodels' library not installed.", template=template)

#     try:
#         df = pd.read_json(io.StringIO(cleaned_json_data[selected_sheet]), orient='split')
        
#         # Prepare data
#         df[date_col] = pd.to_datetime(df[date_col])
#         df = df.sort_values(by=date_col)
        
#         # Aggregate data to monthly sum, as an example
#         ts_data = df.set_index(date_col).resample('ME')[metric_col].sum()
        
#         if len(ts_data) < 24: # Need at least 2 full cycles for seasonal model
#             return go.Figure().update_layout(title=f"Error: Not enough data for forecasting (Need > 24 months, found {len(ts_data)}).", template=template)

#         # Create and fit model
#         model = ExponentialSmoothing(ts_data, seasonal_periods=12, trend='add', seasonal='add').fit()
#         # Forecast 12 months into the future
#         forecast = model.forecast(12)
        
#         # Create plot
#         fig = go.Figure()
#         fig.add_trace(go.Scatter(x=ts_data.index, y=ts_data, mode='lines', name='Historical Data'))
#         fig.add_trace(go.Scatter(x=forecast.index, y=forecast, mode='lines+markers', name='Forecasted Data', line={'dash': 'dash'}))
#         fig.update_layout(title=f"Forecast of {metric_col}", template=template)
#         return fig
        
#     except Exception as e:
#         return go.Figure().update_layout(title=f"Forecast Error: {e}", template=template)


# --- create_dynamic_figure function (no changes) ---
def create_dynamic_figure(chart_type, x_col, y_col, color_col, template, selected_sheet, json_data):
    if not all([chart_type, x_col, json_data, selected_sheet]):
        fig = go.Figure().update_layout(title="Please select chart type and X-axis", template=template)
        return fig
    df = pd.read_json(io.StringIO(json_data[selected_sheet]), orient='split')
    fig = go.Figure()
    is_x_numeric = pd.api.types.is_numeric_dtype(df[x_col])
    is_y_numeric = pd.api.types.is_numeric_dtype(df[y_col]) if y_col else False
    try:
        if chart_type == 'Bar Chart':
            title = f"Bar Chart of {x_col}"
            if not y_col: 
                grouped_df = df.groupby(x_col).size().reset_index(name='Count')
                fig = px.bar(grouped_df, x=x_col, y='Count', color=color_col, title=f"Count of {x_col}", template=template)
            else:
                if not is_y_numeric: return go.Figure().update_layout(title=f"Error: Y-axis ('{y_col}') must be numeric.", template=template)
                title = f"Bar Chart of {y_col} by {x_col}"
                fig = px.bar(df, x=x_col, y=y_col, color=color_col, title=title, template=template)
            fig.update_xaxes(tickangle=45) 
        elif chart_type == 'Line Chart':
            if not y_col: return go.Figure().update_layout(title="Error: Please select a numeric Y-axis.", template=template)
            if not is_y_numeric: return go.Figure().update_layout(title=f"Error: Y-axis ('{y_col}') must be numeric.", template=template)
            title = f"Line Chart of {y_col} by {x_col}"
            fig = px.line(df, x=x_col, y=y_col, color=color_col, title=title, template=template)
            fig.update_xaxes(tickangle=45)
        elif chart_type == 'Scatter Plot':
            if not y_col: return go.Figure().update_layout(title="Error: Please select a numeric Y-axis.", template=template)
            if not is_x_numeric: return go.Figure().update_layout(title=f"Error: X-axis ('{x_col}') must be numeric.", template=template)
            if not is_y_numeric: return go.Figure().update_layout(title=f"Error: Y-axis ('{y_col}') must be numeric.", template=template)
            title = f"Scatter Plot of {y_col} vs {x_col}"
            fig = px.scatter(df, x=x_col, y=y_col, color=color_col, title=title, template=template)
        elif chart_type == 'Histogram':
            if not is_x_numeric: return go.Figure().update_layout(title=f"Error: X-axis ('{x_col}') must be numeric.", template=template)
            title = f"Histogram of {x_col}"
            fig = px.histogram(df, x=x_col, color=color_col, title=title, template=template)
        elif chart_type == 'Pie Chart':
            if not y_col: return go.Figure().update_layout(title="Error: Please select 'Values' (Y-axis).", template=template)
            if is_x_numeric: return go.Figure().update_layout(title=f"Error: 'Names' (X-axis) should be categorical.", template=template)
            if not is_y_numeric: return go.Figure().update_layout(title=f"Error: 'Values' (Y-axis) must be numeric.", template=template)
            title = f"Pie Chart of {y_col} by {x_col} (Names)"
            fig = px.pie(df, names=x_col, values=y_col, color=color_col, title=title, template=template)
    except Exception as e:
        fig = go.Figure().update_layout(title=f"Error creating chart: {e}", template=template)
    return fig

# --- Callbacks for all 6 graphs (now use CLEANED data) ---
@callback(Output('graph1', 'figure'), [Input('chart1-type', 'value'), Input('chart1-x', 'value'), Input('chart1-y', 'value'), Input('chart1-color', 'value'), Input('theme-selector', 'value'), Input('sheet-selector-dropdown', 'value')], [State('cleaned-data-store', 'data')])
def update_graph1(chart_type, x_col, y_col, color_col, template, selected_sheet, json_data): return create_dynamic_figure(chart_type, x_col, y_col, color_col, template, selected_sheet, json_data)
@callback(Output('graph2', 'figure'), [Input('chart2-type', 'value'), Input('chart2-x', 'value'), Input('chart2-y', 'value'), Input('chart2-color', 'value'), Input('theme-selector', 'value'), Input('sheet-selector-dropdown', 'value')], [State('cleaned-data-store', 'data')])
def update_graph2(chart_type, x_col, y_col, color_col, template, selected_sheet, json_data): return create_dynamic_figure(chart_type, x_col, y_col, color_col, template, selected_sheet, json_data)
@callback(Output('graph3', 'figure'), [Input('chart3-type', 'value'), Input('chart3-x', 'value'), Input('chart3-y', 'value'), Input('chart3-color', 'value'), Input('theme-selector', 'value'), Input('sheet-selector-dropdown', 'value')], [State('cleaned-data-store', 'data')])
def update_graph3(chart_type, x_col, y_col, color_col, template, selected_sheet, json_data): return create_dynamic_figure(chart_type, x_col, y_col, color_col, template, selected_sheet, json_data)
@callback(Output('graph4', 'figure'), [Input('chart4-type', 'value'), Input('chart4-x', 'value'), Input('chart4-y', 'value'), Input('chart4-color', 'value'), Input('theme-selector', 'value'), Input('sheet-selector-dropdown', 'value')], [State('cleaned-data-store', 'data')])
def update_graph4(chart_type, x_col, y_col, color_col, template, selected_sheet, json_data): return create_dynamic_figure(chart_type, x_col, y_col, color_col, template, selected_sheet, json_data)
@callback(Output('graph5', 'figure'), [Input('chart5-type', 'value'), Input('chart5-x', 'value'), Input('chart5-y', 'value'), Input('chart5-color', 'value'), Input('theme-selector', 'value'), Input('sheet-selector-dropdown', 'value')], [State('cleaned-data-store', 'data')])
def update_graph5(chart_type, x_col, y_col, color_col, template, selected_sheet, json_data): return create_dynamic_figure(chart_type, x_col, y_col, color_col, template, selected_sheet, json_data)
@callback(Output('graph6', 'figure'), [Input('chart6-type', 'value'), Input('chart6-x', 'value'), Input('chart6-y', 'value'), Input('chart6-color', 'value'), Input('theme-selector', 'value'), Input('sheet-selector-dropdown', 'value')], [State('cleaned-data-store', 'data')])
def update_graph6(chart_type, x_col, y_col, color_col, template, selected_sheet, json_data): return create_dynamic_figure(chart_type, x_col, y_col, color_col, template, selected_sheet, json_data)

# --- Client-side Callback for PDF Printing ---
clientside_callback(
    """function(n_clicks) { if (n_clicks > 0) { window.print(); } return dash.no_update; }""",
    Output('btn-print-pdf', 'n_clicks_timestamp'), 
    Input('btn-print-pdf', 'n_clicks'),
    prevent_initial_call=True
)

# --- MODIFIED: Callback for Excel Export (now with Toast and cleaning) ---
@callback(
    [Output('download-excel', 'data'),
     Output('excel-toast', 'is_open')],
    [Input('btn-export-excel', 'n_clicks')],
    [State('stored-data-summary', 'data'),
     State('cleaned-data-store', 'data')], # Use CLEANED data for export
    prevent_initial_call=True
)
def download_excel_report(n_clicks, summary, cleaned_json_data):
    if not n_clicks or not cleaned_json_data or not summary:
        return dash.no_update, False # Do not open toast

    try:
        # Rebuild the *cleaned* data
        data_dfs = {sheet: pd.read_json(io.StringIO(cleaned_json_data[sheet]), orient='split') for sheet in cleaned_json_data}
        
        # Generate the report based on the *cleaned* data
        report_string = generate_report(summary, data_dfs)
        
        # Call the exporter function
        excel_buffer = excel_exporter.create_excel_report_in_memory(summary, data_dfs, report_string)
        
        # Send file to user and trigger toast
        return dcc.send_bytes(excel_buffer, "Analysis_Dashboard_Report.xlsx"), True

    except Exception as e:
        print(f"FATAL ERROR IN EXCEL EXPORT: {e}")
        return dash.no_update, False


# --- Run the App ---
if __name__ == '__main__':
    app.run(debug=False, port=8050)
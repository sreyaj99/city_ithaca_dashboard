from dash import html, dcc
import pandas as pd
import plotly.express as px

def layout():
    return html.Div([
        html.Div([
            html.H1('Welcome to the Ithaca District Dashboard', style={'color': '#333', 'text-align': 'left'}),
            html.P('This is a dashboard visualizing water consumption in Ithaca.', style={'text-align': 'margin-left'}),
            html.P('Navigate to the pages using the dropdown on the left to explore specific data visualizations.', style={'text-align': 'left'}),
        ], style={'margin-left': '220px', 'padding': '20px'})
    ])
if __name__ == '__main__':
    from dash import Dash
    app = Dash(__name__)
    app.layout = layout()
    app.run_server(debug=True)

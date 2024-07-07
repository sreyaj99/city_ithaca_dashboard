from dash import html, dcc
import pandas as pd
import plotly.express as px
def layout():
    return html.Div([
        html.Div([
            html.H2("2030 Districts", style={'margin-top': '20px'}),
            html.P("2030 Districts are organizations led by the private sector, with local building industry leaders uniting around a shared vision for sustainability and economic growth – while aligning with local community groups and government to achieve significant energy, water, and emissions reductions within our commercial cores. Property owner/manager/developers join a local 2030 District to help them make significant changes to their properties to create reductions necessary to transition to a low carbon economy.", className="rectangle-div")
        ], style={'margin-left': '220px', 'padding': '20px'})
    ])
if __name__ == '__main__':
    from dash import Dash
    app = Dash(__name__)
    app.layout = layout()
    app.run_server(debug=True)
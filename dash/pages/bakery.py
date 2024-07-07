from dash import html, dcc
import pandas as pd
import plotly.express as px

def layout():
    def get_water_line():
        df_water = pd.read_csv('water_bakery.csv')
        df_water['startDate'] = pd.to_datetime(df_water['startDate'])
        fig_water = px.line(df_water, x='startDate', y='usage', title='Water Usage over Time')
        return fig_water

    def get_energy_line():
        df_energy = pd.read_csv('energy_bakery.csv')
        df_energy['startDate'] = pd.to_datetime(df_energy['startDate'])
        fig_energy = px.line(
            df_energy, x='startDate', y='usage', color='ENERGY SOURCE', title='Energy Usage over Time',
            log_y=True  # Enable logarithmic scale on the y-axis
        )
        return fig_energy

    def get_energy_pie():
        df_energy = pd.read_csv('energy_bakery.csv')
        energy_source_counts = df_energy['ENERGY SOURCE'].value_counts()
        fig_energy_pie = px.pie(names=energy_source_counts.index, values=energy_source_counts.values, title='Energy Source Distribution')
        return fig_energy_pie

    return html.Div([
        html.H3('Ithaca Bakery Usage Statistics'),
        
        html.Div([
            html.H4('Water Consumption over Time'),
            dcc.Graph(id='water-graph', figure=get_water_line()),
        ]),

        html.Div([
            html.H4('Energy Consumption over Time'),
            dcc.Graph(id='energy-graph', figure=get_energy_line()),
        ]),

        html.Div([
            html.H4('Energy Source Distribution 2014-2024'),
            dcc.Graph(id='energy-pie-chart', figure=get_energy_pie()),
        ]),

    ], style={'padding': '20px'})

if __name__ == '__main__':
    from dash import Dash
    app = Dash(__name__)
    app.layout = layout()
    app.run_server(debug=True)

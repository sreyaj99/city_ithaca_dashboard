from dash import html, dcc
import pandas as pd
import plotly.express as px

def layout():
    # Create the figure for water consumption
    def get_water_line():
        df_water = pd.read_csv('water_cityhall.csv')
        df_water['startDate'] = pd.to_datetime(df_water['startDate'])
        fig_water = px.line(df_water, x='startDate', y='usage', title='Water Usage over Time')
        return fig_water

    # Create the figure for energy consumption as a line graph
    def get_energy_line():
        df_energy = pd.read_csv('energy_cityhall.csv')
        df_energy['startDate'] = pd.to_datetime(df_energy['startDate'])
        fig_energy = px.line(df_energy, x='startDate', y='usage', color='ENERGY SOURCE', title='Energy Usage over Time')
        return fig_energy

    # Create the figure for energy consumption as a pie chart
    def get_energy_pie():
        df_energy = pd.read_csv('energy_cityhall.csv')
        energy_source_counts = df_energy['ENERGY SOURCE'].value_counts()
        fig_energy_pie = px.pie(names=energy_source_counts.index, values=energy_source_counts.values, title='Energy Source Distribution')
        return fig_energy_pie

    return html.Div([
        html.H3('Cityhall Usage Statistics'),
        
        # Water consumption line graph
        html.Div([
            html.H4('Water Consumption over Time'),
            dcc.Graph(id='water-graph', figure=get_water_line()),
        ]),

        # Energy consumption line graph
        html.Div([
            html.H4('Energy Consumption over Time'),
            dcc.Graph(id='energy-graph', figure=get_energy_line()),
        ]),

        # Energy consumption pie chart
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

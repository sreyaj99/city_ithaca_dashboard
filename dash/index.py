from dash import Dash, html, dcc, Input, Output
from dash.dependencies import State

# Assuming the import of your page modules is correct
import pages.books as books
import pages.bakery as bakery
import pages.cityhall as cityhall
import pages.statetheatre as statetheatre
import pages.glossary as glossary
import pages.home as home

app = Dash(__name__, suppress_callback_exceptions=True)

# Sidebar function for navigation
def sidebar():
    return html.Div([
        html.Img(src="assets/2023.png", style={'width': '100%', 'height': 'auto', 'max-width': '180px', 'padding': '10px'}), 
        html.H2('Navigation', style={'color': '#41ac49', 'text-align': 'center'}),
        dcc.Dropdown(
            id='nav-dropdown',
            options=[
                {'label': 'Home', 'value': '/home'},
                {'label': 'Books', 'value': '/books'},
                {'label': 'Bakery', 'value': '/bakery'},
                {'label': 'City Hall', 'value': '/cityhall'},
                {'label': 'State Theatre', 'value': '/statetheatre'},
                {'label': 'Glossary', 'value': '/glossary'}
            ],
            value='/',
            clearable=False,
            style={'width': '90%', 'margin': '10px auto'}
        )
    ], style={
        'background-color': '#E5E4E2',
        'color': '#333',
        'width': '150px',
        'height': '100vh',
        'position': 'fixed',
        'top': 0,
        'left': 0,
        'padding': '10px'
    })


# Main layout of the app, which includes navigation and content
app.layout = html.Div([
    dcc.Location(id='url', refresh=False),
    html.Div(id='page-content')  # Content will be rendered here based on the URL
])

# Callback to control page content based on URL
@app.callback(Output('page-content', 'children'),
              Input('url', 'pathname'))
def display_page(pathname):
    if pathname == '/books':
        return html.Div([sidebar(), books.layout()], style={'margin-left': '220px', 'padding': '20px'})
    elif pathname == '/bakery':
        return html.Div([sidebar(), bakery.layout()], style={'margin-left': '220px', 'padding': '20px'})
    elif pathname == '/cityhall':
        return html.Div([sidebar(), cityhall.layout()], style={'margin-left': '220px', 'padding': '20px'})
    elif pathname == '/statetheatre':
        return html.Div([sidebar(), statetheatre.layout()], style={'margin-left': '220px', 'padding': '20px'})
    elif pathname == '/home':
        return html.Div([sidebar(), home.layout()], style={'margin-left': '220px', 'padding': '20px'})
    elif pathname == '/glossary':
        return html.Div([sidebar(), glossary.layout()], style={'margin-left': '220px', 'padding': '20px'})
    else:
        return html.Div([
            html.H1("404 Page Not Found", style={'text-align': 'center'}),
            html.P("The page you're looking for doesn't exist.")
        ], style={'margin-left': '220px', 'padding': '20px'})

# Callback to navigate based on dropdown selection
@app.callback(
    Output('url', 'pathname'),
    Input('nav-dropdown', 'value')
)
def update_url(pathname):
    return pathname

if __name__ == '__main__':
    app.run_server(debug=True)

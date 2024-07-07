from dash import Dash
from index import app

# Check if the run statement is under the main guard
if __name__ == '__main__':
    app.run_server(debug=True)

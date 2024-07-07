import xml.etree.ElementTree as ET
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth

# Your Energy Star Portfolio Manager credentials
username = 'Ithaca2030_dashboard'
password = '1/2by2050'

# The Meter ID for which you want to retrieve water meter usage data
meter_id = "36262358"  # Replace with your meter ID

# The API endpoint for retrieving consumption data for a specific meter
api_url = f"https://portfoliomanager.energystar.gov/ws/meter/{meter_id}/consumptionData"

# Make the API request
response = requests.get(api_url, auth=HTTPBasicAuth(username, password))

# Check if the request was successful
if response.status_code == 200:
    print("Successfully retrieved data.")
    
    # Parse the XML response
    root = ET.fromstring(response.content)
    print(response.content)
    # Create an empty list to hold all meter consumption records
    records = []
    
    # Loop through each meterConsumption in the response and extract details
    for consumption in root.findall('.//meterConsumption'):
        record = {
            'startDate': consumption.find('startDate').text if consumption.find('startDate') is not None else None,
            'endDate': consumption.find('endDate').text if consumption.find('endDate') is not None else None,
            'usage': consumption.find('usage').text if consumption.find('usage') is not None else None,
            'cost': consumption.find('cost').text if consumption.find('cost') is not None else None,
            'meterConsumption': consumption.find("meterData").text if consumption.find("meterData") is not None else None
        }
        records.append(record)

    # Convert records to a pandas DataFrame
    df = pd.DataFrame(records)

    # Convert 'usage' and 'cost' to appropriate types (float)
    df['usage'] = pd.to_numeric(df['usage'], errors='coerce')
    df['cost'] = pd.to_numeric(df['cost'], errors='coerce')
    
    # Optionally, convert 'startDate' and 'endDate' to datetime objects
    df['startDate'] = pd.to_datetime(df['startDate'], errors='coerce')
    df['endDate'] = pd.to_datetime(df['endDate'], errors='coerce')

    print(df)
    
    # Save DataFrame to CSV
   # df.to_csv('water_consumption.csv', index=False)
    
   # print("Data saved to water_consumption.csv")
  #  df.to_csv('water_consumption_cityhall.csv', index=False)
    
else:
    print("Failed to retrieve data.", response.status_code) 
import xml.etree.ElementTree as ET
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth

# Your Energy Star Portfolio Manager credentials
username = 'Ithaca2030_dashboard'
password = '1/2by2050'

# The Meter IDs for which you want to retrieve water meter usage data
meter_ids = ["8448417", "8448416"]

# Create an empty list to hold all meter consumption records
all_records = []

# Define a dictionary to map meter IDs to energy sources
energy_sources = {
    "8448417": "Electric",
    "8448416": "Natural Gas",
}

# Loop over each meter ID
for meter_id in meter_ids:
    # The API endpoint for retrieving consumption data for a specific meter
    api_url = f"https://portfoliomanager.energystar.gov/ws/meter/{meter_id}/consumptionData"

    # Make the API request
    response = requests.get(api_url, auth=HTTPBasicAuth(username, password))

    # Check if the request was successful
    if response.status_code == 200:
        print(f"Successfully retrieved data for meter ID: {meter_id}")

        # Parse the XML response
        root = ET.fromstring(response.content)

        # Loop through each meterConsumption in the response and extract details
        for consumption in root.findall('.//meterConsumption'):
            record = {
                'meter_id': meter_id,
                'ENERGY SOURCE': energy_sources[meter_id],  # Assign energy source based on meter ID
                'startDate': consumption.find('startDate').text if consumption.find('startDate') is not None else None,
                'endDate': consumption.find('endDate').text if consumption.find('endDate') is not None else None,
                'usage': consumption.find('usage').text if consumption.find('usage') is not None else None,
                'cost': consumption.find('cost').text if consumption.find('cost') is not None else None
            }
            all_records.append(record)
    else:
        print(f"Failed to retrieve data for meter ID: {meter_id}")

# Convert records to a pandas DataFrame
df = pd.DataFrame(all_records)

# Convert 'usage' and 'cost' to appropriate types (float)
df['usage'] = pd.to_numeric(df['usage'], errors='coerce')
df['cost'] = pd.to_numeric(df['cost'], errors='coerce')

# Optionally, convert 'startDate' and 'endDate' to datetime objects
df['startDate'] = pd.to_datetime(df['startDate'], errors='coerce')
df['endDate'] = pd.to_datetime(df['endDate'], errors='coerce')

print(df)

# Save DataFrame to CSV
df.to_csv('energy_apicall_ cityhall.csv', index=False)

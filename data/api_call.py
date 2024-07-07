import requests
from requests.auth import HTTPBasicAuth
import xml.etree.ElementTree as ET

# Your Energy Star Portfolio Manager credentials
username = 'Ithaca2030_dashboard'
password = '1/2by2050'

 
# The Property Use ID for which you want to retrieve water meter usage data
property_use_id = "4512292"

# The API endpoint for retrieving meter data for a specific property use
api_url = f"https://portfoliomanager.energystar.gov/ws/association/property/{property_use_id}/meter"

# Make the API request
response = requests.get(api_url, auth=HTTPBasicAuth(username, password))

# Check if the request was successful
if response.status_code == 200:
    print("Successfully retrieved data.")
    
# Parse the XML response
root = ET.fromstring(response.content)
print(response.content)

# Find the waterMeterAssociation section
water_meter_association = root.find('.//waterMeterAssociation')

# If a waterMeterAssociation is found
if water_meter_association is not None:
    # Loop through each meterId in the waterMeterAssociation section
    for meter_id in water_meter_association.findall('.//meterId'):
        print(f"Water Meter ID: {meter_id.text}")

        # Now, you could make another API call to get the detailed consumption
        # data for each water meter ID if needed.
else:
    print("No water meter association found.")


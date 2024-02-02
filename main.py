# All Imports
import requests
import pandas as pd
from datetime import datetime
import io

# URL of the zip file
url = "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_240_tg.txt"

# Send HTTP request to the URL
response = requests.get(url)

# Check if the request was successful
if response.status_code == 200:
    # Get the content of the response
    data = io.BytesIO(response.content)

    # Read the txt file into a DataFrame
    df = pd.read_csv(url, skiprows=13, sep=',', low_memory=False)

    # get the current date and time
    now = datetime.now()

    # Format the current date and time as a string
    current_day_str = now.strftime('%Y-%m')

    # Create a filename based on the current date
    filename = f"{current_day_str}.xlsx"

    # Save the DataFrame to an Excel file
    df.to_excel(filename, index=False)
    
    # Print the filename
    print(f"DataFrame has been saved to {filename}")
else:
    # Print an error message if the request was not successful
    print(f"Failed to download the file: {response.status_code}")

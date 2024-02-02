import tkinter as tk
import requests
import pandas as pd
from datetime import datetime
import io
import os 

def download_data(selected_option, message_label):
    # URLs of the txt files
    urls = {
        "Schiphol": "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_240_tg.txt",
        "De Bilt": "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_260_tg.txt"
    }

    # Get the URL based on the selected option
    url = urls[selected_option]

    # Send HTTP request to the URL
    response = requests.get(url)

    # Check if the request was successful
    if response.status_code == 200:
        # Read the txt file into a DataFrame
        df = pd.read_csv(io.StringIO(response.text), skiprows=13, sep=',', low_memory=False)

        # Get the current date
        now = datetime.now()

        # Format the current date
        current_month_str = now.strftime('%Y-%m')

        # Get the current user's home directory
        home_directory = os.path.expanduser('~')

        # Construct the path to the Downloads folder
        downloads_path = os.path.join(home_directory, 'Downloads')

        # Create a filename based on the selected option, current date, and directory
        filename = os.path.join(downloads_path, f"{selected_option}_{current_month_str}.xlsx")

        # Save the DataFrame to an Excel file
        df.to_excel(filename, index=False)
        
        # Update the message label
        message = f"DataFrame has been saved as {selected_option}_{current_month_str}.xlsx at Downloads folder."
        message_label.config(text=message)
    else:
        # Update the message label with an error
        message = f"Failed to download the file: {response.status_code}"
        message_label.config(text=message)

# Create a new Tkinter window
window = tk.Tk()
window.state('zoomed')

# Create a list of options
options = ["Schiphol", "De Bilt"]

# Create a StringVar to hold the selected option
selected_option = tk.StringVar(window)
selected_option.set(options[0])  # Set default value

# Create the OptionMenu and pack it
select_box = tk.OptionMenu(window, selected_option, *options)
select_box.pack()

# Create a Label for displaying messages
message_label = tk.Label(window, text="", anchor="e")
message_label.pack(fill=tk.X, side=tk.BOTTOM)

# Create a new button
button = tk.Button(window, text="Download", command=lambda: download_data(selected_option.get(), message_label))

# Add the button to the window
button.pack()

# Start the Tkinter event loop
window.mainloop()

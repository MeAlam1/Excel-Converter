import tkinter as tk
import requests
import pandas as pd
from datetime import datetime
import io
import os

def download_data(selected_option, message_label):
    urls = {
        "Den Helder/De Kooy": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_235_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_235_txg.txt"
        ],
        "Leeuwarden": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_270_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_270_txg.txt"
        ],
        "Groningen/Eelde": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_280_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_280_txg.txt"
        ],
        "Twenthe": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_290_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_290_txg.txt"
        ],
        "Schiphol": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_240_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_240_txg.txt"
        ],
        "De Bilt": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_260_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_260_txg.txt"
        ],
        "Rotterdam": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_344_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_344_txg.txt"
        ],
        "Vlissingen": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_310_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_310_txg.txt"
        ],
        "Eindhoven": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_370_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_370_txg.txt"
        ],
        "Maastricht/Beek": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_380_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_380_txg.txt"
        ],
    }

    selected_urls = urls[selected_option]
    now = datetime.now()
    current_month_str = now.strftime('%Y-%m')
    home_directory = os.path.expanduser('~')
    downloads_path = os.path.join(home_directory, 'Downloads')
    filename = os.path.join(downloads_path, f"{selected_option.replace('/', '_')}_{current_month_str}.xlsx")

    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        for i, url in enumerate(selected_urls, start=1):
            response = requests.get(url)
            if response.status_code == 200:
                df = pd.read_csv(io.StringIO(response.content.decode('utf-8')), skiprows=13, sep=',', low_memory=False)
                sheet_name = f"Sheet_{i}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                message = f"Failed to download from URL {i}: {response.status_code}"
                message_label.config(text=message)
                return

    message = f"Data has been saved as {filename} in the Downloads folder."
    message_label.config(text=message)

window = tk.Tk()
window.state('zoomed')
options = ["Den Helder/De Kooy", "Leeuwarden", "Groningen/Eelde", "Twenthe", "Schiphol", "De Bilt", "Rotterdam", "Vlissingen", "Eindhoven", "Maastricht/Beek"]
selected_option = tk.StringVar(window)
selected_option.set(options[0])
select_box = tk.OptionMenu(window, selected_option, *options)
select_box.pack()
message_label = tk.Label(window, text="", anchor="e")
message_label.pack(fill=tk.X, side=tk.BOTTOM)
button = tk.Button(window, text="Download", command=lambda: download_data(selected_option.get(), message_label))
button.pack()
window.mainloop()

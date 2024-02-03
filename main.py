import tkinter as tk
from tkinter import ttk
import threading
import requests
import pandas as pd
from datetime import datetime
import io
import os

def download_data(selected_option, message_label, separate_folder=False):
    urls = {
        "Den Helder/De Kooy": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_235_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_235_txg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_235_tng.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_235_rh24.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_235_pg.txt"
        ],
        "Leeuwarden": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_270_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_270_txg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_270_tng.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_270_rh24.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_270_pg.txt"
        ],
        "Groningen/Eelde": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_280_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_280_txg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_280_tng.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_280_rh24.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_280_pg.txt"
        ],
        "Twenthe": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_290_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_290_txg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_290_tng.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_290_rh24.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_290_pg.txt"
        ],
        "Schiphol": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_240_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_240_txg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_240_tng.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_240_rh24.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_240_pg.txt"
        ],
        "De Bilt": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_260_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_260_txg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_260_tng.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_260_rh24.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_260_pg.txt"
        ],
        "Rotterdam": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_344_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_344_txg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_344_tng.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_344_rh24.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_344_pg.txt"
        ],
        "Vlissingen": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_310_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_310_txg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_310_tng.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_310_rh24.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_310_pg.txt"
        ],
        "Eindhoven": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_370_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_370_txg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_370_tng.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_370_rh24.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_370_pg.txt"
        ],
        "Maastricht/Beek": [
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_380_tg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_380_txg.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_380_tng.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_380_rh24.txt",
            "https://cdn.knmi.nl/knmi/map/page/klimatologie/gegevens/maandgegevens/mndgeg_380_pg.txt"
        ],
    }

    custom_sheet_names = {
        "tg": "Gemiddelde temperatuur",
        "txg": "Gemiddelde maximumtemperatuur",
        "tng": "Gemiddelde minimumtemperatuur",
        "rh24": "Neerslagsom",
        "pg": "Gemiddelde luchtdruk"
    }

    selected_urls = urls[selected_option]
    now = datetime.now()
    current_month_str = now.strftime('%Y-%m')
    home_directory = os.path.expanduser('~')
    downloads_path = os.path.join(home_directory, 'Downloads')
    if separate_folder:
        new_folder_path = os.path.join(downloads_path, f"{current_month_str}_All_Locations")
        os.makedirs(new_folder_path, exist_ok=True)
        filename = os.path.join(new_folder_path, f"{selected_option.replace('/', '_')}_{current_month_str}.xlsx")
    else:
        filename = os.path.join(downloads_path, f"{selected_option.replace('/', '_')}_{current_month_str}.xlsx")
   

    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        for url in selected_urls:
            response = requests.get(url)
            if response.status_code == 200:
                url_part = url.split('_')[-1].split('.')[0] 
                custom_name = custom_sheet_names.get(url_part, "Sheet")

                df = pd.read_csv(io.StringIO(response.content.decode('utf-8')), skiprows=13, sep=',', low_memory=False)
                df.to_excel(writer, sheet_name=custom_name, index=False)
            else:
                message = f"Failed to download from URL: {response.status_code}"
                message_label.config(text=message)
                return

    message = f"Data has been saved as {selected_option.replace('/', '_')}_{current_month_str} in the Downloads folder."
    message_label.config(text=message)

def download_all_data(options, message_label, progress_bar):
    def thread_target():
        total = len(options)
        for i, option in enumerate(options):
            download_data(option, message_label, separate_folder=True)
            progress = (i + 1) / total * 100
            progress_bar['value'] = progress
            window.update_idletasks()
        message_label.config(text="All data has been downloaded.")
        progress_bar['value'] = 0

    threading.Thread(target=thread_target).start()


window = tk.Tk()
window.title("KNMI Data Converter")
window.iconbitmap('knmi.ico')
window.state('zoomed')


style = ttk.Style()
style.theme_use('clam')
style.configure('TButton', font=('Arial', 10, 'bold'), background='#EEE')
style.configure('TOptionMenu', font=('Arial', 10, 'bold'), background='#EEE')

options = ["Den Helder/De Kooy", "Leeuwarden", "Groningen/Eelde", "Twenthe", "Schiphol", "De Bilt", "Rotterdam", "Vlissingen", "Eindhoven", "Maastricht/Beek"]
selected_option = tk.StringVar(window)
selected_option.set(options[0])

select_box = ttk.OptionMenu(window, selected_option, *options)
select_box.pack()

download_button = ttk.Button(window, text="Download Selected Location", command=lambda: download_data(selected_option.get(), message_label))
download_button.pack()

download_all_button = ttk.Button(window, text="Download All Locations", command=lambda: download_all_data(options, message_label, progress_bar))
download_all_button.pack()

progress_bar = ttk.Progressbar(window, orient=tk.HORIZONTAL, length=100, mode='determinate')
progress_bar.pack()

message_label = tk.Label(window, text="", anchor="e")
message_label.pack()

window.mainloop()
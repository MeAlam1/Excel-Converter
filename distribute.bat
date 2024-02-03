@echo off
pyinstaller --onefile --name FileConverter --icon=knmi.ico --noconsole --add-data "knmi.ico;." main.py


TMS Data Processing
20 Mei 2025

pyinstaller --onefile --noconsole --windowed --icon=icon.ico --name="TMS Data Processing" --strip --add-data "icon.ico;." --add-data "constant.json;." --add-data "secret.json;." --add-data "modules;modules" --add-data "utils;utils" --hidden-import "babel.numbers" --upx-dir="upx" apps.py
rd /s /q build
rd /s /q dist
del /f main.spec
C:\Users\PC\anaconda3\envs\weatherApp\Scripts\pyinstaller.exe -F -w -i "radio.ico" --add-data "config.ini;." --add-data "data;data/" --hidden-import "pandas" --hidden-import "pandas.core.arrays.arrow" --hidden-import "openpyxl" --hidden-import "lxml" --hidden-import "bs4" main.py

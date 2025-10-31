@echo off
cd /d "%~dp0"
call env\Scripts\activate
start C:\laragon\laragon.exe
start http://127.0.0.1:5000
python app.py
pause
@echo off
cd /d "%~dp0"
set CNAMIS_PORT=8766
".venv\Scripts\python.exe" app.py

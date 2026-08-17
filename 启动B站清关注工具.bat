@echo off
cd /d "%~dp0"
pythonw bili_follow_manager_gui.py
if errorlevel 1 python bili_follow_manager_gui.py
pause

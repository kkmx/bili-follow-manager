@echo off
chcp 65001 >nul
python -m pip install pyinstaller
pyinstaller --onefile --windowed --name BiliFollowManager --add-data "partition_map.json;." bili_follow_manager_gui.py
echo.
echo 打包完成：dist\BiliFollowManager.exe
pause

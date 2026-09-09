@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PY=C:\Users\Administrator\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

echo.
echo 正在把两份 CSV 转换成 announcement.json ...
echo.

"%PY%" build_json.py

echo.
echo 完成后把 announcement.json 推到 GitHub 即生效（游戏端几分钟内刷新）。
echo.
pause

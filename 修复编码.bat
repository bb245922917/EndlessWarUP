@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PY=C:\Users\Administrator\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

echo.
echo 把公告 CSV 转成 UTF-8，并恢复被 Excel 存坏的韩文
echo （用之前请先关闭 Excel / WPS，否则文件被占用会失败）
echo.

"%PY%" build_json.py --fix

echo.
pause

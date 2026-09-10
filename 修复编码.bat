@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PY=C:\Users\Administrator\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
if not exist "%PY%" set "PY=C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\python.exe"

echo.
echo 检查并修复 XLSX 里被破坏的韩文（若出现 ?）
echo （XLSX 本身是 Unicode 存储，正常情况下不会坏；只有当文件曾被当成 CSV 存过才可能丢韩文）
echo （用之前请先关闭 Excel / WPS，否则文件被占用会失败）
echo.

"%PY%" build_json.py --fix

echo.
pause

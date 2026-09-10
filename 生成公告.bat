@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PY=C:\Users\Administrator\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
if not exist "%PY%" set "PY=C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\python.exe"

echo ===== 打印表格全部数据（核对用，不生成文件）=====
"%PY%" build_json.py --dump
echo.
echo ===== 生成 announcement.json =====
"%PY%" build_json.py
if errorlevel 1 (
  echo.
  echo [失败] Python 脚本未正常结束，请把上面的报错截图发我
)
pause

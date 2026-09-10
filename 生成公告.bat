@echo off
cd /d "D:\Project\Unreal\无尽战争公告更新"
"C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\python.exe" build_json.py
if errorlevel 1 (
  echo.
  echo [失败] Python 脚本未正常结束，请把上面的报错截图发我
)
pause

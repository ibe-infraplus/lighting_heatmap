@echo off
setlocal
cd /d "%~dp0"
py generate_lighting_dashboard.py --excel test.xlsx --output lighting_dashboard.html
if errorlevel 1 (
  echo.
  echo สร้าง Dashboard ไม่สำเร็จ กรุณาตรวจสอบข้อความด้านบน
  pause
  exit /b 1
)
echo.
echo เปิด Local Server ที่ http://localhost:8000/lighting_dashboard.html
start "" http://localhost:8000/lighting_dashboard.html
py -m http.server 8000

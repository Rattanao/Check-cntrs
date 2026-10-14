@echo off
setlocal
cd /d "%~dp0"
echo ============================================
echo   Check CNTRS - ตรวจเทียบ MANIFEST กับ CNTRS
echo ============================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [X] ไม่พบ Python กรุณาติดตั้งจาก https://www.python.org/downloads/
    echo     อย่าลืมติ๊ก "Add python.exe to PATH" ตอนติดตั้ง
    echo.
    pause
    exit /b 1
)

echo กำลังตรวจสอบไลบรารีที่ต้องใช้ ...
python -m pip install -q -r requirements.txt
if errorlevel 1 (
    echo [X] ติดตั้งไลบรารีไม่สำเร็จ ตรวจสอบอินเทอร์เน็ตแล้วลองใหม่
    echo.
    pause
    exit /b 1
)

echo.
python check_cntrs.py

echo.
pause

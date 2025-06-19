@echo off
echo Auto Consolidator - Quick Build Script
echo =====================================
echo.

echo Installing required packages...
pip install -r requirements.txt
echo.

echo Building executable...
python build_executable.py
echo.

echo Build process complete!
pause

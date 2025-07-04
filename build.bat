@echo off
REM Auto Consolidator Build Script
REM This script builds the executable using PyInstaller

echo Building Auto Consolidator executable...
echo.

REM Check if PyInstaller is installed
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo PyInstaller not found. Installing...
    pip install pyinstaller
    echo.
)

REM Clean previous build
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

REM Build executable
echo Creating executable...
REM Check if icon file exists and use appropriate build method
if exist "auto_consolidator.ico" (
    echo Using custom icon: auto_consolidator.ico
    pyinstaller Auto_Consolidator.spec
) else (
    echo No icon file found. Building without custom icon...
    echo See icon_instructions.txt for how to add a custom icon.
    pyinstaller --onefile --windowed --name "Auto_Consolidator" auto_consolidator.py
)

if exist "dist\Auto_Consolidator.exe" (
    echo.
    echo ========================================
    echo Build completed successfully!
    echo Executable location: dist\Auto_Consolidator.exe
    echo ========================================
    echo.
    
    REM Copy Cell Map template to dist folder
    if exist "Cell Map.xlsx" (
        copy "Cell Map.xlsx" "dist\"
        echo Cell Map template copied to dist folder.
    )
    
    echo Opening dist folder...
    start "" "dist"
) else (
    echo.
    echo ========================================
    echo Build failed! Check the output above for errors.
    echo ========================================
)

pause
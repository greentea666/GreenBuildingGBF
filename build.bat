@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo  Building GreenBuildingGBF.exe ...
echo.
"C:\Users\User\AppData\Local\Programs\Python\Python314\python.exe" -m PyInstaller ^
    --onefile --windowed ^
    --name "GreenBuildingGBF" ^
    --icon "app.ico" ^
    --add-data "app.ico;." ^
    --noconfirm ^
    green_app.py
echo.
echo  Done! EXE is in: dist\GreenBuildingGBF.exe
echo.
pause

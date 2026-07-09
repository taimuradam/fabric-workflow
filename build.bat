@echo off
REM ============================================================
REM  Build TextileCosting.exe  —  double-click or run from cmd
REM  Requires the venv to be set up first (see README).
REM ============================================================

call venv\Scripts\activate

pyinstaller --clean --onefile --windowed --noconsole --name TextileCosting --icon assets\icon.ico --add-data "assets;assets" --collect-all openpyxl --collect-all customtkinter main.py

echo.
echo ============================================================
echo  Done. Your app is at:  dist\TextileCosting.exe
echo.
echo  If Explorer still shows the OLD icon, that's Windows'
echo  icon cache - rename or copy the .exe and it will refresh.
echo ============================================================
pause

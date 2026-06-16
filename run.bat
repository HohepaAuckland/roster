@echo off
echo ================================================
echo   HOHEPA AUCKLAND ROSTER SYSTEM
echo ================================================
echo.
echo Checking Python...
python --version 2>nul || (echo Python not found. Please install Python 3 from python.org && pause && exit)
echo.
echo Installing required packages...
pip install flask openpyxl pandas xlsxwriter --quiet
echo.
echo Starting Roster App...
echo.
echo Open your browser and go to:
echo   http://localhost:5050
echo.
echo Press Ctrl+C to stop the app.
echo ================================================
python app.py
pause

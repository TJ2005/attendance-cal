@echo off
echo ========================================
echo Attendance Calculator Setup and Run
echo ========================================
echo.

REM Check if virtual environment exists
if not exist ".venv" (
    echo Creating virtual environment...
    python -m venv .venv
    echo Virtual environment created.
    echo.
)

REM Activate virtual environment
echo Activating virtual environment...
call .venv\Scripts\activate.bat

REM Install requirements
echo.
echo Installing dependencies...
pip install -r requirements.txt

REM Run the attendance calculator
echo.
echo ========================================
echo Running Attendance Calculator...
echo ========================================
echo.
python attendance_calculator.py

REM Open the HTML report in default browser
if exist "attendance_report.html" (
    echo.
    echo Opening report in browser...
    start attendance_report.html
)

echo.
echo ========================================
echo Done!
echo ========================================
pause

@echo off
setlocal
echo === MarkItDown Setup ===
echo.
echo [1/2] Creating virtual environment...
python -m venv venv
if errorlevel 1 goto :fail

call venv\Scripts\activate.bat
if errorlevel 1 goto :fail

echo [2/2] Installing packages...
python -m pip install --upgrade pip
if errorlevel 1 goto :fail

python -m pip install -r requirements.txt
if errorlevel 1 goto :fail

echo.
echo Setup complete! Run the app with:
echo   venv\Scripts\activate
echo   python markitdown_app.py
echo.
pause
exit /b 0

:fail
echo.
echo Setup failed.
pause
exit /b 1

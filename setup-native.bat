@echo off
setlocal

echo === MarkItDown Native Setup ===
echo.

set "HAS_DOTNET_SDK="
for /f %%I in ('dotnet --list-sdks 2^>nul') do set "HAS_DOTNET_SDK=1"
if not defined HAS_DOTNET_SDK (
    echo ERROR: .NET SDK not found.
    echo Install the .NET 8 SDK first:
    echo   https://dotnet.microsoft.com/en-us/download/dotnet/8.0
    echo.
    pause
    exit /b 1
)

python --version >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python was not found in PATH.
    echo Install Python 3.11 or newer, then run this script again.
    echo.
    pause
    exit /b 1
)

if not exist "venv\Scripts\python.exe" (
    echo [1/3] Creating virtual environment...
    python -m venv venv
    if errorlevel 1 goto :fail
) else (
    echo [1/3] Reusing existing virtual environment...
)

echo [2/3] Activating virtual environment...
call "venv\Scripts\activate.bat"
if errorlevel 1 goto :fail

echo [3/3] Installing Python dependencies for the temporary fallback backend...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 goto :fail

echo.
echo Setup complete.
echo The app now uses native C# converters first, with Python kept only as a fallback during migration.
echo Open the native app project in:
echo   Native\MarkItDown.Native.sln
echo Or build it with:
echo   build-native.bat
echo.
pause
exit /b 0

:fail
echo.
echo Setup failed.
pause
exit /b 1

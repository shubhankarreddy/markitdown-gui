@echo off
setlocal

echo === MarkItDown Build ===
echo.

if not exist "venv\Scripts\activate.bat" (
    echo ERROR: Virtual environment not found.
    echo Run setup.bat first.
    echo.
    pause
    exit /b 1
)

echo [1/3] Activating virtual environment...
call "venv\Scripts\activate.bat"
if errorlevel 1 goto :fail

tasklist /FI "IMAGENAME eq MarkItDown.exe" | find /I "MarkItDown.exe" >nul
if not errorlevel 1 (
    echo ERROR: MarkItDown.exe is still running.
    echo Close the app and run build.bat again so PyInstaller can replace dist\MarkItDown\MarkItDown.exe.
    echo.
    pause
    exit /b 1
)

echo [2/3] Building executable...
if exist "dist\MarkItDown.exe" del /Q "dist\MarkItDown.exe" >nul 2>nul
python -m PyInstaller --noconfirm "MarkItDown.spec"
if errorlevel 1 goto :fail

set "ISCC_PATH="
for %%I in (ISCC.exe) do if not defined ISCC_PATH set "ISCC_PATH=%%~$PATH:I"
if not defined ISCC_PATH if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC_PATH=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC_PATH if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC_PATH=%ProgramFiles%\Inno Setup 6\ISCC.exe"

if defined ISCC_PATH (
    echo [3/3] Building Windows installer...
    "%ISCC_PATH%" "installer.iss"
    if errorlevel 1 goto :fail
    echo.
    echo Build complete.
    echo   dist\MarkItDown\MarkItDown.exe
    echo   installer_output\MarkItDown_Setup.exe
) else (
    echo [3/3] Inno Setup Compiler was not found.
    echo Skipping installer build. Install Inno Setup 6 to generate MarkItDown_Setup.exe.
    echo.
    echo Executable is ready at dist\MarkItDown\MarkItDown.exe
)

echo.
pause
exit /b 0

:fail
echo.
echo Build failed.
pause
exit /b 1

@echo off
setlocal

echo === MarkItDown Native Build ===
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

if not exist "venv\Scripts\python.exe" (
    echo ERROR: Python virtual environment not found.
    echo Run setup-native.bat first.
    echo.
    pause
    exit /b 1
)

echo [1/4] Activating virtual environment...
call "venv\Scripts\activate.bat"
if errorlevel 1 goto :fail

echo [2/4] Building Python fallback helper...
python -m PyInstaller --noconfirm "MarkItDownBackend.spec"
if errorlevel 1 goto :fail

set "PUBLISH_DIR=%CD%\native_publish\win-x64"
echo [3/4] Publishing WPF app...
dotnet publish "Native\MarkItDown.Native\MarkItDown.Native.csproj" -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:IncludeNativeLibrariesForSelfExtract=true -o "%PUBLISH_DIR%"
if errorlevel 1 goto :fail

copy /Y "dist\MarkItDownBackend.exe" "%PUBLISH_DIR%\MarkItDownBackend.exe" >nul
if errorlevel 1 goto :fail

set "ISCC_PATH="
for %%I in (ISCC.exe) do if not defined ISCC_PATH set "ISCC_PATH=%%~$PATH:I"
if not defined ISCC_PATH if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC_PATH=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC_PATH if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC_PATH=%ProgramFiles%\Inno Setup 6\ISCC.exe"

if defined ISCC_PATH (
    echo [4/4] Building Windows installer...
    "%ISCC_PATH%" "installer-native.iss"
    if errorlevel 1 goto :fail
    echo.
    echo Build complete.
    echo   %PUBLISH_DIR%\MarkItDown.exe
    echo   installer_output\MarkItDown_Native_Setup.exe
) else (
    echo [4/4] Inno Setup Compiler was not found.
    echo Skipping installer build.
    echo.
    echo Build complete.
    echo   %PUBLISH_DIR%\MarkItDown.exe
)

echo.
pause
exit /b 0

:fail
echo.
echo Build failed.
pause
exit /b 1

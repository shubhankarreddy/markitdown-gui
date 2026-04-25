# MarkItDown Installer
# Run: Right-click -> "Run with PowerShell"
# Or:  powershell -ExecutionPolicy Bypass -File install.ps1

$ErrorActionPreference = "Stop"

$appName = "MarkItDown"
$exeName = "MarkItDown.exe"
$srcExe  = Join-Path $PSScriptRoot "dist\$exeName"

if (-not (Test-Path $srcExe)) {
    Write-Host ""
    Write-Host "ERROR: dist\MarkItDown.exe not found." -ForegroundColor Red
    Write-Host "Build it first by running setup.bat then build.bat" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

$installDir = Join-Path $env:LOCALAPPDATA $appName
New-Item -ItemType Directory -Path $installDir -Force | Out-Null

Copy-Item $srcExe -Destination $installDir -Force
$installedExe = Join-Path $installDir $exeName
Write-Host "Installed $exeName to $installDir" -ForegroundColor Green

# Desktop shortcut
$ws = New-Object -ComObject WScript.Shell
$desktopLink = Join-Path ([Environment]::GetFolderPath("Desktop")) "$appName.lnk"
$sc = $ws.CreateShortcut($desktopLink)
$sc.TargetPath = $installedExe
$sc.WorkingDirectory = $installDir
$sc.Description = "Document to Markdown converter"
$sc.Save()
Write-Host "Created Desktop shortcut" -ForegroundColor Green

# Start Menu shortcut
$startMenu = Join-Path ([Environment]::GetFolderPath("StartMenu")) "Programs"
$startLink = Join-Path $startMenu "$appName.lnk"
$sc2 = $ws.CreateShortcut($startLink)
$sc2.TargetPath = $installedExe
$sc2.WorkingDirectory = $installDir
$sc2.Description = "Document to Markdown converter"
$sc2.Save()
Write-Host "Created Start Menu shortcut" -ForegroundColor Green

# Create uninstaller
$uninstallScript = Join-Path $installDir "uninstall.ps1"
$uninstallContent = @"
Remove-Item `"$desktopLink`" -Force -ErrorAction SilentlyContinue
Remove-Item `"$startLink`" -Force -ErrorAction SilentlyContinue
Remove-Item `"$installDir`" -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "MarkItDown has been uninstalled." -ForegroundColor Green
Read-Host "Press Enter to exit"
"@
Set-Content -Path $uninstallScript -Value $uninstallContent
Write-Host "Created uninstaller" -ForegroundColor Green

Write-Host ""
Write-Host "=== Installation complete\! ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Launch MarkItDown from:"
Write-Host "  - Desktop shortcut"
Write-Host "  - Start Menu -> MarkItDown"
Write-Host ""
Write-Host "To uninstall later, run:"
Write-Host "  $uninstallScript"
Write-Host ""
Read-Host "Press Enter to exit"

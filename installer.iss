; MarkItDown Installer — Inno Setup Script
; Produces a proper Windows installer with Start Menu, Desktop shortcut, and uninstaller.
;
; Prerequisites:
;   1. Build the exe first:  build.bat
;   2. Install Inno Setup:   https://jrsoftware.org/isinfo.php
;   3. Right-click this file -> "Compile" (or open in Inno Setup Compiler)
;
; The installer will be created in the "Output" folder.

[Setup]
AppId={{D1353E3A-1B5F-4BD8-A5F5-50A0A77C0C2C}
AppName=MarkItDown
AppVersion=1.0.0
AppVerName=MarkItDown 1.0.0
AppPublisher=MarkItDown
DefaultDirName={localappdata}\Programs\MarkItDown
DefaultGroupName=MarkItDown
UninstallDisplayIcon={app}\MarkItDown.exe
OutputDir=installer_output
OutputBaseFilename=MarkItDown_Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
WizardResizable=no
PrivilegesRequired=lowest
DisableProgramGroupPage=yes
UsePreviousAppDir=yes
UsePreviousGroup=yes
CloseApplications=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "dist\MarkItDown\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\MarkItDown";         Filename: "{app}\MarkItDown.exe"
Name: "{autodesktop}\MarkItDown";   Filename: "{app}\MarkItDown.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\MarkItDown.exe"; Description: "Launch MarkItDown"; Flags: nowait postinstall skipifsilent

; MarkItDown Native Installer - Inno Setup Script

[Setup]
AppId={{8A4307E9-34B8-4EE4-9010-845A31A3EF31}
AppName=MarkItDown
AppVersion=2.0.0
AppVerName=MarkItDown Native 2.0.0
AppPublisher=MarkItDown
DefaultDirName={localappdata}\Programs\MarkItDown
DefaultGroupName=MarkItDown
UninstallDisplayIcon={app}\MarkItDown.exe
OutputDir=installer_output
OutputBaseFilename=MarkItDown_Native_Setup
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
Source: "native_publish\win-x64\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\MarkItDown"; Filename: "{app}\MarkItDown.exe"
Name: "{autodesktop}\MarkItDown"; Filename: "{app}\MarkItDown.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\MarkItDown.exe"; Description: "Launch MarkItDown"; Flags: nowait postinstall skipifsilent

; Inno Setup script for AmanInventory

[Setup]
AppName=AmanInventory
AppVersion=1.0.0
DefaultDirName={pf}\AmanInventory
DefaultGroupName=AmanInventory
OutputDir=installer\output
OutputBaseFilename=AmanInventorySetup
Compression=lzma
SolidCompression=yes

[Files]
Source: "..\dist\AmanInventory.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\assets\*"; DestDir: "{app}\assets"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\db_config.json"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\AmanInventory"; Filename: "{app}\AmanInventory.exe"
Name: "{commondesktop}\AmanInventory"; Filename: "{app}\AmanInventory.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop icon"; GroupDescription: "Additional icons:"

[Run]
Filename: "{app}\AmanInventory.exe"; Description: "Launch AmanInventory"; Flags: nowait postinstall skipifsilent

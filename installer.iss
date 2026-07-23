[Setup]
AppId={{ED260038-C62C-41D8-BBEA-ED3E5F0D3787}}
AppName=Report Generator
AppVersion=1.0.5
AppPublisher=Your Name
DefaultDirName={autopf}\ReportGenerator
DefaultGroupName=Report Generator
DisableProgramGroupPage=yes
OutputDir=installer_output
OutputBaseFilename=ReportGenerator-Setup-1.0.5
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\report-generator.exe

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "dist\report-generator\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Report Generator"; Filename: "{app}\report-generator.exe"
Name: "{group}\Uninstall Report Generator"; Filename: "{uninstallexe}"
Name: "{commondesktop}\Report Generator"; Filename: "{app}\report-generator.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"

[Run]
Filename: "{app}\report-generator.exe"; Description: "Launch Report Generator"; Flags: postinstall nowait skipifsilent

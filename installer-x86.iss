[Setup]
AppName=Outlook Classic Search
AppVersion=1.2.0
AppPublisher=mjeppo
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
DefaultDirName={autopf}\Outlook Classic Search
DefaultGroupName=Outlook Classic Search
OutputDir=installer-output
OutputBaseFilename=OutlookClassicSearch-v1.2.0-Setup-x86
SetupIconFile=ocs_icon_transparant02.ico
Compression=lzma2/ultra64
SolidCompression=yes
ArchitecturesAllowed=x86 x64
ArchitecturesInstallIn64BitMode=
WizardStyle=modern
MinVersion=10.0
PrivilegesRequired=lowest
UninstallDisplayIcon={app}\OutlookClassicSearch.exe

[Languages]
Name: "dutch"; MessagesFile: "compiler:Languages\Dutch.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Snelkoppeling op het bureaublad"; GroupDescription: "Extra opties:"; Flags: unchecked

[Files]
Source: "bin\Release\net8.0-windows\win-x86\publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{userprograms}\Outlook Classic Search"; Filename: "{app}\OutlookClassicSearch.exe"
Name: "{userprograms}\Verwijderen"; Filename: "{uninstallexe}"
Name: "{userdesktop}\Outlook Classic Search"; Filename: "{app}\OutlookClassicSearch.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\OutlookClassicSearch.exe"; Description: "Outlook Classic Search starten"; Flags: nowait postinstall skipifsilent

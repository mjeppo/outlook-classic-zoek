[Setup]
AppName=Outlook Classic Search
AppVersion=1.3.0
AppPublisher=mjeppo
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
DefaultDirName={autopf}\Outlook Classic Search
DefaultGroupName=Outlook Classic Search
OutputDir=installer-output
OutputBaseFilename=OutlookClassicSearch-v1.3.0-Setup-x86
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
; Installeer .NET 8 Desktop Runtime als die nog niet aanwezig is
Filename: "{tmp}\dotnet-installer.exe"; Parameters: "/install /quiet /norestart"; \
  StatusMsg: ".NET 8 Desktop Runtime installeren..."; \
  Check: NeedsDotNet8; Flags: waituntilterminated

Filename: "{app}\OutlookClassicSearch.exe"; Description: "Outlook Classic Search starten"; Flags: nowait postinstall skipifsilent

[Code]
// Controleer of .NET 8.x Desktop Runtime (x86) aanwezig is
function NeedsDotNet8: Boolean;
var
  keyBase: String;
  i: Integer;
  subKey: String;
begin
  Result := True;
  keyBase := 'SOFTWARE\dotnet\Setup\InstalledVersions\x86\sharedfx\Microsoft.WindowsDesktop.App';
  i := 0;
  while RegEnumKeyEx(HKLM, keyBase, i, subKey, 0) do
  begin
    if Copy(subKey, 1, 2) = '8.' then
    begin
      Result := False;
      Exit;
    end;
    Inc(i);
  end;
end;

procedure InitializeWizard;
var
  url: String;
  dest: String;
begin
  if NeedsDotNet8 then
  begin
    url := 'https://builds.dotnet.microsoft.com/dotnet/WindowsDesktop/8.0/latest/windowsdesktop-runtime-win-x86.exe';
    dest := ExpandConstant('{tmp}\dotnet-installer.exe');
    idpAddFile(url, dest);
    idpDownloadAfter(wpReady);
  end;
end;

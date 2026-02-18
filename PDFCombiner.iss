
[Setup]
AppId={{6F2D6A6E-6E94-4C13-8E2B-0D50C7A5F0F2}}
AppName=PDFCombiner
AppVersion=2.1.2
OutputBaseFilename=PDFCombinerInstaller_2.1.2
AppPublisher=tgtechy
DefaultDirName={autopf}\PDFCombiner
DefaultGroupName=PDFCombiner
OutputDir=installer
Compression=lzma
SolidCompression=yes
SetupIconFile=pdfcombinericon.ico
WizardStyle=modern

[Files]
Source: "dist\PDFCombiner.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "LICENSE"; DestDir: "{app}"; Flags: ignoreversion
Source: "pdfcombinericon.ico"; DestDir: "{app}"; Flags: ignoreversion
;Source: "instructions.html"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\PDFCombiner"; Filename: "{app}\PDFCombiner.exe"; IconFilename: "{app}\pdfcombinericon.ico"
;Name: "{group}\Instructions"; Filename: "{app}\instructions.html"

[Run]
Filename: "{app}\PDFCombiner.exe"; Description: "Launch PDFCombiner"; Flags: nowait postinstall skipifsilent

[Code]

var
  DeleteConfig: Boolean;

procedure InitializeUninstallProgressForm();
begin
  if MsgBox('Do you want to remove your PDFCombiner settings (settings.json) from AppData?', mbConfirmation, MB_YESNO) = IDYES then
    DeleteConfig := True
  else
    DeleteConfig := False;
end;


procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ConfigPath: String;
begin
  if (CurUninstallStep = usUninstall) and DeleteConfig then
  begin
    ConfigPath := ExpandConstant('{userappdata}\\PDFCombiner\\settings.json');
    if FileExists(ConfigPath) then
      DeleteFile(ConfigPath);
  end;

  // After uninstall, try to remove the install directory if empty
  if CurUninstallStep = usUninstall then
  begin
    RemoveDir(ExpandConstant('{app}'));
  end;
end;
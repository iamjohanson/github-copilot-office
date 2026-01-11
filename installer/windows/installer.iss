; Inno Setup Script for GitHub Copilot Office Add-in
; Requires Inno Setup 6.x: https://jrsoftware.org/isinfo.php

#define MyAppName "GitHub Copilot Office Add-in"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Your Company"
#define MyAppURL "https://github.com/your-org/copilot-sdk-office-sample"
#define MyAppExeName "copilot-office-server.exe"

[Setup]
AppId={{12345678-1234-1234-1234-123456789012}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=..\..\build\windows
OutputBaseFilename=CopilotOfficeAddin-Setup-{#MyAppVersion}
SetupIconFile=app.ico
UninstallDisplayIcon={app}\app.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "startatlogin"; Description: "Start automatically when Windows starts"; GroupDescription: "Additional options:"

[Files]
; Main executable (built with pkg)
Source: "..\..\build\windows\copilot-office-server.exe"; DestDir: "{app}"; Flags: ignoreversion

; Application icon
Source: "app.ico"; DestDir: "{app}"; Flags: ignoreversion

; Static files
Source: "..\..\dist\*"; DestDir: "{app}\dist"; Flags: ignoreversion recursesubdirs createallsubdirs

; Certificates
Source: "..\..\certs\*"; DestDir: "{app}\certs"; Flags: ignoreversion

; Manifest
Source: "..\..\manifest.xml"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"

[Registry]
; Register the add-in manifest with Office
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\Developer"; ValueType: string; ValueName: "CopilotOfficeAddin"; ValueData: "{app}\manifest.xml"; Flags: uninsdeletevalue

[Run]
; Trust the SSL certificate
Filename: "certutil.exe"; Parameters: "-addstore -user Root ""{app}\certs\localhost.pem"""; Flags: runhidden; StatusMsg: "Installing SSL certificate..."

; Start the service after install
Filename: "{app}\{#MyAppExeName}"; Description: "Start {#MyAppName} now"; Flags: nowait postinstall skipifsilent

[UninstallRun]
; Stop the service before uninstall
Filename: "taskkill.exe"; Parameters: "/F /IM {#MyAppExeName}"; Flags: runhidden

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Code]
var
  StartupTaskCreated: Boolean;

procedure CreateStartupTask();
var
  TaskXML: string;
  TaskFile: string;
  ResultCode: Integer;
begin
  TaskXML := '<?xml version="1.0" encoding="UTF-16"?>' + #13#10 +
    '<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">' + #13#10 +
    '  <Triggers>' + #13#10 +
    '    <LogonTrigger>' + #13#10 +
    '      <Enabled>true</Enabled>' + #13#10 +
    '    </LogonTrigger>' + #13#10 +
    '  </Triggers>' + #13#10 +
    '  <Principals>' + #13#10 +
    '    <Principal id="Author">' + #13#10 +
    '      <LogonType>InteractiveToken</LogonType>' + #13#10 +
    '      <RunLevel>LeastPrivilege</RunLevel>' + #13#10 +
    '    </Principal>' + #13#10 +
    '  </Principals>' + #13#10 +
    '  <Settings>' + #13#10 +
    '    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>' + #13#10 +
    '    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>' + #13#10 +
    '    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>' + #13#10 +
    '    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>' + #13#10 +
    '    <Hidden>true</Hidden>' + #13#10 +
    '  </Settings>' + #13#10 +
    '  <Actions Context="Author">' + #13#10 +
    '    <Exec>' + #13#10 +
    '      <Command>"' + ExpandConstant('{app}\{#MyAppExeName}') + '"</Command>' + #13#10 +
    '      <WorkingDirectory>' + ExpandConstant('{app}') + '</WorkingDirectory>' + #13#10 +
    '    </Exec>' + #13#10 +
    '  </Actions>' + #13#10 +
    '</Task>';
  
  TaskFile := ExpandConstant('{tmp}\CopilotOfficeTask.xml');
  SaveStringToFile(TaskFile, TaskXML, False);
  
  Exec('schtasks.exe', '/Create /TN "CopilotOfficeAddin" /XML "' + TaskFile + '" /F', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  StartupTaskCreated := (ResultCode = 0);
end;

procedure RemoveStartupTask();
var
  ResultCode: Integer;
begin
  Exec('schtasks.exe', '/Delete /TN "CopilotOfficeAddin" /F', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    if IsTaskSelected('startatlogin') then
      CreateStartupTask();
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    RemoveStartupTask();
  end;
end;

; ExcelVerifier Installer Script
; Requires Inno Setup 6.0 or later

#define MyAppName "ExcelVerifier"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "ExcelVerifier"
#define MyAppExeName "ExcelVerifier.exe"

[Setup]
AppId={{8B3D4F2A-9C1E-4D5F-A7B8-1E2F3C4D5E6F}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=installer_output
OutputBaseFilename=ExcelVerifier_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "polish"; MessagesFile: "compiler:Languages\Polish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "ExcelVerifier\dist\ExcelVerifier\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "ExcelVerifier\icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\configure_api.exe"; Parameters: "--api-key ""{code:GetApiKey}"""; Flags: runhidden; Description: "Configure API Key"; StatusMsg: "Configuring API key..."
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
var
  ApiKeyPage: TInputQueryWizardPage;

procedure InitializeWizard;
begin
  { Create API Key input page }
  ApiKeyPage := CreateInputQueryPage(wpSelectTasks,
    'Google Gemini API Key', 
    'Enter your API key for image transformation feature',
    'Please enter your Google Gemini API key. You can obtain one from https://aistudio.google.com/app/apikey' + #13#10 + 
    'If you skip this step, you can configure it later in Settings.');
  
  ApiKeyPage.Add('API Key:', False);
end;

function GetApiKey(Param: String): String;
begin
  Result := ApiKeyPage.Values[0];
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := False;
end;

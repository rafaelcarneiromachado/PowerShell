; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "THOTH"
#define MyAppVersion "1.0"
#define MyAppPublisher "Rafael Machado"
#define MyAppURL "https://github.com/rafaelcarneiromachado"

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{B031001F-907F-46AD-AF22-48B3D1A582A7}
;SignTool=signtool
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={sd}\LogCollector\SetupFiles
DisableDirPage=yes
DisableFinishedPage=yes
DisableReadyPage=yes
DisableWelcomePage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
; Uncomment the following line to run in non administrative install mode (install for current user only.)
;PrivilegesRequired=lowest
OutputDir=D:\PSScripts\PowerShell\THOTH\
OutputBaseFilename=THOTH
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
Uninstallable=no
AppCopyright=OpenSource
ArchitecturesAllowed=x86 x64
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
ButtonNext=&Run >

[Files]
Source: "D:\PSScripts\PowerShell\THOTH\BIN\Thoth-Form.ps1"; DestDir: "{app}"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Run]
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-ExecutionPolicy ByPass -WindowStyle Hidden -NoProfile -file C:\LogCollector\SetupFiles\Thoth-Form.ps1";Flags: nowait

[Code]
#ifdef UNICODE
  #define AW "W"
#else
  #define AW "A"
#endif
type
  HINSTANCE = THandle;

function ShellExecute(hwnd: HWND; lpOperation: string; lpFile: string;
  lpParameters: string; lpDirectory: string; nShowCmd: Integer): HINSTANCE;
  external 'ShellExecute{#AW}@shell32.dll stdcall';

function InitializeSetup: Boolean;
begin
  // if this instance of the setup is not silent which is by running
  // setup binary without /SILENT parameter, stop the initialization
  Result := WizardSilent;
  // if this instance is not silent, then...
  if not Result then
  begin
    // re-run the setup with /SILENT parameter; because executing of
    // the setup loader is not possible with ShellExec function, we
    // need to use a WinAPI workaround
    if ShellExecute(0, '', ExpandConstant('{srcexe}'), '/SILENT', '',
      SW_SHOW) <= 32
    then
      // if re-running this setup to silent mode failed, let's allow
      // this non-silent setup to be run
      Result := True;
  end;
end;

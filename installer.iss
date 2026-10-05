#define MyAppName "Noor-ul-Hifz"
#define MyAppVersion "2.0.0"
#define MyAppPublisher "ALI QURESHI"
#define MyAppExeName "Noor-ul-Hifz.exe"

[Setup]
AppId={{A7C6B2F1-8B47-4E5E-9B5D-NOORULHIFZ20}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

DefaultDirName={autopf}\Noor-ul-Hifz
DefaultGroupName={#MyAppName}

OutputDir=release
OutputBaseFilename=Noor-ul-Hifz-Setup-2.0


Compression=lzma
SolidCompression=yes

WizardStyle=modern

PrivilegesRequired=admin

ArchitecturesInstallIn64BitMode=x64compatible

DisableProgramGroupPage=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "dist\Noor-ul-Hifz.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Noor-ul-Hifz"; Filename: "{app}\Noor-ul-Hifz.exe"
Name: "{group}\Uninstall Noor-ul-Hifz"; Filename: "{uninstallexe}"

Name: "{autodesktop}\Noor-ul-Hifz"; Filename: "{app}\Noor-ul-Hifz.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Noor-ul-Hifz.exe"; Description: "Launch Noor-ul-Hifz"; Flags: nowait postinstall skipifsilent
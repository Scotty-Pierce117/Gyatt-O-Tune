; Gyatt-O-Tune Inno Setup Installer Script
; Requires: Inno Setup 6+ (https://jrsoftware.org/isinfo.php)
;
; The portable EXE must be built first:
;   powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
;
; Then compile this script:
;   iscc installer.iss
; Or let build_exe.ps1 do both automatically if ISCC.exe is on your PATH
; or installed at the default Inno Setup location.

#define AppName      "Gyatt-O-Tune"
#define AppVersion   "0.1.2"
#define AppPublisher "Gyatt-O-Tune"
#define AppURL       "https://github.com/gyatt-o-tune"
#define AppExeName   "Gyatt-O-Tune.exe"
#define AppExeSource "dist\Gyatt-O-Tune.exe"
#define AppIcon      "src\gyatt_o_tune\assets\gyatt-o-tune.ico"
#define WizardLarge  "src\gyatt_o_tune\assets\installer-wizard-large.bmp"
#define WizardSmall  "src\gyatt_o_tune\assets\installer-wizard-small.bmp"

[Setup]
; Stable AppId GUID — do NOT change after first release or it breaks upgrade detection
AppId={{6F3DA82C-1E4B-4A7F-9C55-2B8EFA034D71}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}

; Install per-user by default (no UAC), but let the user choose All Users at the prompt
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

; Default install location: per-user → %LocalAppData%\Programs\Gyatt-O-Tune
;                           all-users → %ProgramFiles%\Gyatt-O-Tune
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes

; Output
OutputDir=dist
OutputBaseFilename=Gyatt-O-Tune-Setup-{#AppVersion}
SetupIconFile={#AppIcon}

; Compression
Compression=lzma2/ultra64
SolidCompression=yes
LZMAUseSeparateProcess=yes

WizardStyle=modern
DisableWelcomePage=no
DisableProgramGroupPage=auto
WizardImageFile={#WizardLarge}
WizardSmallImageFile={#WizardSmall}

; Require Windows 10 or later (same as PySide6 minimum)
MinVersion=10.0.17763

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "fileassoc"; Description: "Associate .msq tune files with {#AppName}"; GroupDescription: "File associations:"; Flags: unchecked

[Files]
; The single portable EXE — everything is bundled inside it by PyInstaller
Source: "{#AppExeSource}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}";                         Filename: "{app}\{#AppExeName}"
Name: "{group}\{cm:UninstallProgram,{#AppName}}";   Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}";                   Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; \
    Description: "{cm:LaunchProgram,{#AppName}}"; \
    Flags: nowait postinstall skipifsilent

[Registry]
; ProgID — created when file association task is checked
Root: HKA; Subkey: "Software\Classes\GyattOTune.MSQFile";                                          ValueType: string; ValueName: "";        ValueData: "MegaSquirt Tune File";               Flags: uninsdeletekey;   Tasks: fileassoc
Root: HKA; Subkey: "Software\Classes\GyattOTune.MSQFile\DefaultIcon";                              ValueType: string; ValueName: "";        ValueData: "{app}\{#AppExeName},0";              Flags: uninsdeletekey;   Tasks: fileassoc
Root: HKA; Subkey: "Software\Classes\GyattOTune.MSQFile\shell\open\command";                       ValueType: string; ValueName: "";        ValueData: """{app}\{#AppExeName}"" ""%1""";     Flags: uninsdeletekey;   Tasks: fileassoc
; Map .msq extension to our ProgID (only when task is checked)
Root: HKA; Subkey: "Software\Classes\.msq";                                                        ValueType: string; ValueName: "";        ValueData: "GyattOTune.MSQFile";                 Flags: uninsdeletevalue; Tasks: fileassoc
; Always register in OpenWithProgids so "Open with" shows the app even without full association
Root: HKA; Subkey: "Software\Classes\.msq\OpenWithProgids";                                        ValueType: string; ValueName: "GyattOTune.MSQFile"; ValueData: "";                        Flags: uninsdeletevalue
; Registered Applications entry for the Windows "Default Programs" control panel
Root: HKA; Subkey: "Software\{#AppName}\Capabilities\FileAssociations";                            ValueType: string; ValueName: ".msq";    ValueData: "GyattOTune.MSQFile";                 Flags: uninsdeletekey
Root: HKA; Subkey: "Software\RegisteredApplications";                                              ValueType: string; ValueName: "{#AppName}"; ValueData: "Software\{#AppName}\Capabilities"; Flags: uninsdeletevalue

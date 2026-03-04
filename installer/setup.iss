; ============================================
;  EDTECH DOC TEMPLATER — Inno Setup Script
;  Compiles into a professional Windows installer EXE
; ============================================

#define MyAppName "EDTECH DOC TEMPLATER"
#define MyAppVersion "3.5"
#define MyAppPublisher "EDTECH Suite"
#define MyAppURL "https://github.com/Israx1990BO/doc-templater"

[Setup]
AppId={{E2D4C8A1-5F3B-4A2E-9C1D-7B8E6F0A3D5E}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename=Setup_EDTECH_DOC_TEMPLATER_v{#MyAppVersion}
SetupIconFile=app_icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
OutputDir=dist

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Messages]
spanish.WelcomeLabel2=Este asistente instalará {#MyAppName} v{#MyAppVersion} en su computadora.%n%nSe instalarán automáticamente:%n  • Python 3.11%n  • Tesseract OCR%n  • Todas las dependencias necesarias%n%nSe requiere conexión a Internet.

[Files]
; Include the PowerShell setup script
Source: "installer\install_deps.ps1"; DestDir: "{tmp}"; Flags: deleteafterinstall
; Include the app icon
Source: "app_icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -WindowStyle Hidden -Command ""cd '{app}'; .\venv\Scripts\Activate.ps1; python test_app.py"""; WorkingDir: "{app}"; IconFilename: "{app}\app_icon.ico"
Name: "{commondesktop}\{#MyAppName}"; Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -WindowStyle Hidden -Command ""cd '{app}'; .\venv\Scripts\Activate.ps1; python test_app.py"""; WorkingDir: "{app}"; IconFilename: "{app}\app_icon.ico"

[Run]
; Run the dependency installer during setup
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{tmp}\install_deps.ps1"" -InstallDir ""{app}"""; StatusMsg: "Instalando dependencias (esto puede tomar unos minutos)..."; Flags: runhidden waituntilterminated

[UninstallDelete]
Type: filesandordirs; Name: "{app}\venv"
Type: filesandordirs; Name: "{app}\uploads"
Type: filesandordirs; Name: "{app}\outputs"
Type: filesandordirs; Name: "{app}\__pycache__"

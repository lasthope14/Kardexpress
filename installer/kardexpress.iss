; Instalador de Kardexpress (Inno Setup 6).
; Se compila desde build.ps1, que pasa la versión:  ISCC.exe /DAppVersion=2.1.0 installer\kardexpress.iss

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#define AppName "Kardexpress"
#define AppPublisher "Hidroobras"
#define AppExe "Kardexpress.exe"

[Setup]
; El AppId identifica la app entre versiones: NO cambiarlo, o las actualizaciones se instalarían aparte.
AppId={{B413E061-0740-4B6A-8E69-EFA18DD841F2}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
VersionInfoVersion={#AppVersion}
VersionInfoCompany={#AppPublisher}
VersionInfoDescription=Instalador de {#AppName}
; Instalación por usuario: no pide permisos de administrador (útil en equipos de oficina)
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
DisableDirPage=yes
OutputDir=..\dist
OutputBaseFilename=Kardexpress-Instalador-{#AppVersion}
SetupIconFile=..\assets\kardexpress.ico
UninstallDisplayIcon={app}\{#AppExe}
UninstallDisplayName={#AppName}
WizardStyle=modern
WizardSizePercent=100
Compression=lzma2/max
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; Si la app está abierta durante una actualización, se cierra y se vuelve a abrir
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "es"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear un acceso directo en el escritorio"; GroupDescription: "Accesos directos:"

[Files]
; Antes de copiar la versión nueva se borra la anterior para no dejar archivos viejos
Source: "..\dist\Kardexpress\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[InstallDelete]
Type: filesandordirs; Name: "{app}\_internal"

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "Abrir {#AppName} ahora"; Flags: nowait postinstall skipifsilent

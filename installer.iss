;---------------------------------------
; Script Inno Setup pour OpenLautrec
; Installation utilisateur (compatible PC scolaires)
;---------------------------------------

[Setup]
AppName=OpenLautrec
AppVersion=1.4.11
PrivilegesRequired=lowest
DefaultDirName={localappdata}\OpenLautrec
DefaultGroupName=OpenLautrec
DisableProgramGroupPage=yes
UsePreviousAppDir=yes
OutputDir=installer
OutputBaseFilename=OpenLautrec-Setup
Compression=lzma
SolidCompression=yes
SetupIconFile=logo.ico
WizardStyle=modern
ChangesAssociations=yes
LicenseFile=LICENSE.txt
AppComments=Traitement de texte - Lycée Toulouse-Lautrec
AppContact=kasperweis23@gmail.com

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Files]
Source: "main.dist\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion
Source: "logo.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{userprograms}\OpenLautrec"; Filename: "{app}\main.exe"; IconFilename: "{app}\logo.ico"; WorkingDir: "{app}"

Name: "{userdesktop}\OpenLautrec"; Filename: "{app}\main.exe"; IconFilename: "{app}\logo.ico"; WorkingDir: "{app}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Créer un raccourci sur le bureau"; GroupDescription: "Options supplémentaires"
Name: "associate"; Description: "Associer l'extension .olc à OpenLautrec"; GroupDescription: "Options supplémentaires"
Name: "contextmenu"; Description: "Ajouter 'Ouvrir avec OpenLautrec' au menu clic droit"; GroupDescription: "Intégration Windows"
Name: "autostart"; Description: "Lancer OpenLautrec au démarrage de Windows"; GroupDescription: "Démarrage"; Flags: unchecked

[Run]
Filename: "{app}\main.exe"; Description: "Lancer OpenLautrec"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Registry]

Root: HKCU; Subkey: "Software\Classes\.olc"; ValueType: string; ValueData: "OpenLautrec.Document"; Flags: uninsdeletevalue; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\OpenLautrec.Document"; ValueType: string; ValueData: "Document OpenLautrec"; Flags: uninsdeletekey; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\OpenLautrec.Document\DefaultIcon"; ValueType: string; ValueData: "{app}\main.exe,0"; Tasks: associate
Root: HKCU; Subkey: "Software\Classes\OpenLautrec.Document\shell\open\command"; ValueType: string; ValueData: """{app}\main.exe"" ""%1"""; Tasks: associate

Root: HKCU; Subkey: "Software\Classes\*\shell\OpenLautrec"; ValueType: string; ValueData: "Ouvrir avec OpenLautrec"; Flags: uninsdeletekey; Tasks: contextmenu
Root: HKCU; Subkey: "Software\Classes\*\shell\OpenLautrec"; ValueName: "Icon"; ValueType: string; ValueData: "{app}\main.exe,0"; Tasks: contextmenu
Root: HKCU; Subkey: "Software\Classes\*\shell\OpenLautrec\command"; ValueType: string; ValueData: """{app}\main.exe"" ""%1"""; Tasks: contextmenu

Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "OpenLautrec"; ValueData: """{app}\main.exe"""; Flags: uninsdeletevalue; Tasks: autostart
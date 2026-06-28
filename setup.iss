[Setup]
AppName=MsgViewer
AppVersion=1.0.0
DefaultDirName={autopf}\MsgViewer
DefaultGroupName=MsgViewer
UninstallDisplayIcon={app}\MsgViewer.exe
Compression=lzma2/max
SolidCompression=yes
OutputDir=D:\prject\Appbuild
OutputBaseFilename=MsgViewerSetup
SetupIconFile=D:\prject\app_icon.ico
AppPublisher=xuanhieu
AppPublisherURL=https://github.com/lestmegogo/Viewmsg
DisableProgramGroupPage=yes
PrivilegesRequired=admin
ChangesAssociations=yes

[Files]
Source: "D:\prject\Appbuild\MsgViewer.exe"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
Root: HKA; Subkey: "Software\Classes\.msg"; ValueType: string; ValueName: ""; ValueData: "MsgViewer.Document"; Flags: uninsdeletevalue
Root: HKA; Subkey: "Software\Classes\MsgViewer.Document"; ValueType: string; ValueName: ""; ValueData: "Outlook Message File"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\MsgViewer.Document\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\MsgViewer.exe,0"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\MsgViewer.Document\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\MsgViewer.exe"" ""%1"""; Flags: uninsdeletekey

Root: HKA; Subkey: "Software\Classes\.eml"; ValueType: string; ValueName: ""; ValueData: "MsgViewer.Document.Eml"; Flags: uninsdeletevalue
Root: HKA; Subkey: "Software\Classes\MsgViewer.Document.Eml"; ValueType: string; ValueName: ""; ValueData: "MIME Email File"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\MsgViewer.Document.Eml\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\MsgViewer.exe,0"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\MsgViewer.Document.Eml\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\MsgViewer.exe"" ""%1"""; Flags: uninsdeletekey

[Icons]
Name: "{group}\MsgViewer"; Filename: "{app}\MsgViewer.exe"
Name: "{autodesktop}\MsgViewer"; Filename: "{app}\MsgViewer.exe"

[Run]
Filename: "{app}\MsgViewer.exe"; Description: "Launch MsgViewer"; Flags: postinstall nowait skipifsilent

[Setup]
AppName=MsgViewer
AppVersion=1.0.0
DefaultDirName={autopf}\MsgViewer
DefaultGroupName=MsgViewer
UninstallDisplayIcon={app}\MsgViewer.exe
Compression=lzma2
SolidCompression=yes
OutputDir=D:\prject\Appbuild
OutputBaseFilename=setup
AppPublisher=xuanhieu
AppSupportURL=https://github.com/lestmegogo/Viewmsg
WizardStyle=modern

[Files]
Source: "D:\prject\bin\Release\net8.0-windows\publish_self\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs

[Icons]
Name: "{group}\MsgViewer"; Filename: "{app}\MsgViewer.exe"
Name: "{autodesktop}\MsgViewer"; Filename: "{app}\MsgViewer.exe"

[Run]
Filename: "{app}\MsgViewer.exe"; Description: "Launch MsgViewer"; Flags: postinstall nowait skipifsilent

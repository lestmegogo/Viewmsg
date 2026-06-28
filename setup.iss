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

[Files]
Source: "D:\prject\Appbuild\MsgViewer.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\MsgViewer"; Filename: "{app}\MsgViewer.exe"
Name: "{autodesktop}\MsgViewer"; Filename: "{app}\MsgViewer.exe"

[Run]
Filename: "{app}\MsgViewer.exe"; Description: "Launch MsgViewer"; Flags: postinstall nowait skipifsilent

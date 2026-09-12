[Setup]
AppName=RaSahLembur
AppVersion=1.1
DefaultDirName={pf}\RaSahLembur
DefaultGroupName=RaSahLembur
OutputBaseFilename=RaSahLemburInstaller
Compression=lzma
SolidCompression=yes

[Files]
Source: "dist\RaSahLembur.exe"; DestDir: "{app}"; Flags: ignoreversion
; Optionally include VC++ redistributable
; Source: "vcredist_x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Run]
; Install VC++ runtime silently if included
; Filename: "{tmp}\vcredist_x64.exe"; Parameters: "/install /quiet /norestart"; StatusMsg: "Installing runtime..."; Flags: runhidden
Filename: "{app}\RaSahLembur.exe"; Description: "Launch RaSahLembur"; Flags: nowait postinstall skipifsilent

[Setup]
AppName=Exportador ETABS
AppVersion=1.0
DefaultDirName={commonpf}\ExportadorETABS
OutputBaseFilename=ExportadorETABS-Setup
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible

[Files]
Source: "dist\gui_tablas_etabs.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Exportador ETABS"; Filename: "{app}\gui_tablas_etabs.exe"
Name: "{commondesktop}\Exportador ETABS"; Filename: "{app}\gui_tablas_etabs.exe"

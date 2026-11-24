[Setup]
AppId=LocalDocxToPdfConverter
AppName=DOCX to PDF Converter
AppVersion=1.0.1
DefaultDirName={autopf}\DOCX to PDF Converter
DisableProgramGroupPage=yes
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=dist
OutputBaseFilename=DOCX-to-PDF-Setup
Compression=lzma2
SolidCompression=yes
UninstallDisplayName=DOCX to PDF Converter

[Files]
Source: "ConvertToPDF.ps1"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
Root: HKLM; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\LocalDocxToPdf"; ValueType: string; ValueData: "Konwertuj do PDF"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\LocalDocxToPdf"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Single"
Root: HKLM; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\LocalDocxToPdf\command"; ValueType: string; ValueData: """{sys}\WindowsPowerShell\v1.0\powershell.exe"" -NoProfile -STA -WindowStyle Hidden -File ""{app}\ConvertToPDF.ps1"" ""%1"""
Root: HKLM; Subkey: "Software\Classes\SystemFileAssociations\.doc\shell\LocalDocxToPdf"; ValueType: string; ValueData: "Konwertuj do PDF"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Classes\SystemFileAssociations\.doc\shell\LocalDocxToPdf"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Single"
Root: HKLM; Subkey: "Software\Classes\SystemFileAssociations\.doc\shell\LocalDocxToPdf\command"; ValueType: string; ValueData: """{sys}\WindowsPowerShell\v1.0\powershell.exe"" -NoProfile -STA -WindowStyle Hidden -File ""{app}\ConvertToPDF.ps1"" ""%1"""

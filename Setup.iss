[Setup]
AppId=LocalDocxToPdfConverter
AppName=DOCX to PDF Converter
AppPublisher=Kevin Fedak
VersionInfoCompany=Kevin Fedak
AppVersion=2.1.0-beta.1
VersionInfoVersion=2.1.0.0
DefaultDirName={autopf}\DOCX to PDF Converter
DisableProgramGroupPage=yes
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=dist
OutputBaseFilename=DOCX-to-PDF-Setup-2.1.0-beta.1
Compression=lzma2
SolidCompression=yes
UninstallDisplayName=DOCX to PDF Converter

[Files]
Source: "WordPdfConverter.exe"; DestDir: "{app}"; Flags: ignoreversion

[InstallDelete]
Type: files; Name: "{app}\Launcher.exe"
Type: files; Name: "{app}\ConvertToPDF.ps1"
[Registry]
Root: HKLM; Subkey: "Software\Classes\SystemFileAssociations\.doc\shell\LocalDocxToPdf"; Flags: deletekey dontcreatekey
Root: HKLM; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\LocalDocxToPdf"; Flags: deletekey dontcreatekey
Root: HKLM; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\KevinWordPdf"; ValueType: string; ValueData: "Konwertuj do PDF (2.1 beta)"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\KevinWordPdf"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Single"
Root: HKLM; Subkey: "Software\Classes\SystemFileAssociations\.docx\shell\KevinWordPdf\command"; ValueType: string; ValueData: """{app}\WordPdfConverter.exe"" ""%1"""
Root: HKLM; Subkey: "Software\Classes\SystemFileAssociations\.doc\shell\KevinWordPdf"; ValueType: string; ValueData: "Konwertuj do PDF (2.1 beta)"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Classes\SystemFileAssociations\.doc\shell\KevinWordPdf"; ValueType: string; ValueName: "MultiSelectModel"; ValueData: "Single"
Root: HKLM; Subkey: "Software\Classes\SystemFileAssociations\.doc\shell\KevinWordPdf\command"; ValueType: string; ValueData: """{app}\WordPdfConverter.exe"" ""%1"""

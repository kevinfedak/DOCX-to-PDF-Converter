# DOCX to PDF Converter

Windows 10/11, desktop Microsoft Word required. Local conversion using Word COM in the interactive user's session. No application telemetry. Word's own network behavior and security policies remain applicable.

## Installation
Download the artifact from Actions, extract ZIP and run the EXE as administrator. On Windows 11 use Show more options. Right-click one DOCX and select Konwertuj do PDF. Existing PDFs are not overwritten. The installer is unsigned. PowerShell execution policy is not bypassed; sign the script if required by policy.

Silent install: `DOCX-to-PDF-Setup.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART`. Uninstall via Installed Apps or the installed uninstaller with the same switches.

## Acceptance tests (not yet executed)
- Install and verify context menu as a standard user.
- Convert a DOCX with spaces and Polish characters; compare layout with Word export.
- Verify existing PDF remains unchanged.
- Verify error handling for missing Word, denied write permissions and invalid documents.
- Verify macro security and enterprise policies; do not use untrusted documents for routine conversion.
- Uninstall and verify menu and installed files are removed.

CI validates PowerShell syntax and compiles the installer; it does not test Word COM.

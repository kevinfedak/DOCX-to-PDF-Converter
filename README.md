# DOCX to PDF Converter 2.1 beta

Windows utility that adds **Konwertuj do PDF (2.1 beta)** to the context menu for DOC and DOCX files. It requires the desktop version of Microsoft Word for Windows and administrator rights to install. Right-click a document and choose the command; the PDF is created next to the original. An existing PDF is never overwritten. On Windows 11, the command may appear under “Show more options”.

Download the test installer from the successful GitHub Actions run's artifacts. The installer replaces the earlier version and updates its context-menu entries. Uninstall it through Windows installed apps.

The build checks installation, registry commands and input validation. The hosted CI runner does not include Word, so actual document conversion, formatting, password-protected documents and Word dialogs still need a manual test on a Windows machine with Word installed. Use copies of representative documents for the beta test.

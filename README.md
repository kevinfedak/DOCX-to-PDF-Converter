# DOCX to PDF Converter 2.1 beta

Windows utility that adds **Konwertuj do PDF (2.1 beta)** to the context menu for DOC and DOCX files. It requires the desktop version of Microsoft Word for Windows and administrator rights to install. Select up to 15 documents in Explorer, right-click and choose the command; each document is converted in turn. For one document, use the same command; the PDF is created next to the original. An existing PDF is never overwritten. On Windows 11, the command may appear under “Show more options”.

Download the test installer from the successful GitHub Actions run's artifacts. The installer replaces the earlier version and updates its context-menu entries. Uninstall it through Windows installed apps.

The build checks installation, registry commands and input validation. The hosted CI runner does not include Word, so actual document conversion, formatting, password-protected documents and Word dialogs still need a manual test on a Windows machine with Word installed. Use copies of representative documents for the beta test.

Explorer starts a separate converter process for each selected file. Conversions run one at a time in a shared Word session and progress window, which lists successes and errors. Larger selections can be converted in batches of up to 15. The installer is not digitally signed; Windows will show an unknown publisher until it is signed with a trusted code-signing certificate.

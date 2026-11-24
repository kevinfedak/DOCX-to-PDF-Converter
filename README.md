# DOCX to PDF Converter

A Windows context-menu utility for converting Word documents to PDF using Microsoft Word.

## Installation

Download `DOCX-to-PDF-Setup-2.1.2.exe` from the [latest release](https://github.com/kevinfedak/DOCX-to-PDF-Converter/releases/latest) and run it. Installation requires administrator rights and desktop Microsoft Word for Windows. The installer updates earlier versions. It is not digitally signed, so Windows displays an unknown publisher warning.

## Use

Select DOC or DOCX files in Explorer, right-click and choose **Konwertuj do PDF**. On Windows 11, the command may appear under **Show more options**. Explorer normally supports up to 100 selected files for this menu entry. Files are processed sequentially in one Word session; the progress window shows results and errors. PDFs are saved next to their source documents. Existing PDFs are never overwritten.

Uninstall the converter through Windows installed apps.

## Building and testing

GitHub Actions builds the executable and installer on Windows and verifies installation, menu registration and input validation. The hosted runner does not include Word, so conversion requires a separate manual test on a Windows machine with Word installed.

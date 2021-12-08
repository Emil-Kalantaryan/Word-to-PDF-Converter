@ECHO OFF
CD /D "%~dp0"
START /MAX "Word to PDF Converter - 1.0.6" Powershell -ExecutionPolicy ByPass -File "Word_to_PDF_Converter.ps1"

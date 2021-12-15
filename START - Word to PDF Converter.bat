@ECHO OFF
CD /D "%~dp0"
START /MAX "Word to PDF Converter" Powershell -ExecutionPolicy ByPass -File "Word_to_PDF_Converter.ps1"

@ECHO OFF
CD /D "%~dp0"
START /MAX Powershell -ExecutionPolicy ByPass -File "Word_to_PDF_Converter.ps1"

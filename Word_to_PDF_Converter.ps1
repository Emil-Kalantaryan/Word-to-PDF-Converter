# Author: 	Emil Kalantaryan
# Name: 	Word to PDF Converter
# Version: 	1.0.0
# Date: 	04/12/2021

# Required folders cheking
if (!(Test-Path -Path ($PSScriptRoot + "\Input"))) {
    $LogLine = "[$(Get-Date)] Creating the folder: '$PSScriptRoot\Input'"
    Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
    Write-Host $LogLine -ForegroundColor "Cyan"
    New-Item -Path $PSScriptRoot -Name "Input" -ItemType "Directory" > $Null
}
if (!(Test-Path -Path ($PSScriptRoot + "\Output"))) {
    $LogLine = "[$(Get-Date)] Creating the folder: '$PSScriptRoot\Output'"
    Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
    Write-Host $LogLine -ForegroundColor "Cyan"
    New-Item -Path $PSScriptRoot -Name "Output" -ItemType "Directory" > $Null
}

# File System Watcher - Declaration & Configuration
$FileSystemWatcher = New-Object System.IO.FileSystemWatcher
$FileSystemWatcher.Path = [System.Environment]::CurrentDirectory + "\Input"
$FileSystemWatcher.Filter = "*.doc*"
$FileSystemWatcher.IncludeSubdirectories = $False
$FileSystemWatcher.EnableRaisingEvents = $True
$FileSystemWatcher | Get-Member -MemberType Event

# Word to PDF Converter - Script
$Convert_WORD_to_PDF = {
	# Current File information
	$FileName = $Event.SourceEventArgs.Name
	$FilePath = $Event.SourceEventArgs.FullPath

	# Word Application Object - Declaration & Configuration
	$Word = New-Object -ComObject Word.Application
	$Word.Visible = $False

	# Converted file Output
	$OutputPath = [System.Environment]::CurrentDirectory + "\Output\" + $FileName.Substring(0,$FileName.LastIndexOf("."))

	$LogLine = "`n[$(Get-Date)] New file detected: $FileName"
	Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
	Write-Host $LogLine -ForegroundColor "Cyan"

	$LogLine = "[$(Get-Date)] Generating PDF of the file: $FileName"
	Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
	Write-Host $LogLine -ForegroundColor "Yellow"

	$Document = $Word.Documents.Open($FilePath)
	$Document.SaveAs([ref] $OutputPath, [ref] 17)
	$Document.Close()

	$LogLine = "[$(Get-Date)] The PDF of the file '$FileName' has been generated successfully."
	Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
	Write-Host $LogLine -ForegroundColor "Green"

	$LogLine = "[$(Get-Date)] Deleting the file: $FileName"
	Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
	Write-Host $LogLine -ForegroundColor "Yellow"

	Remove-Item $FilePath

	$LogLine = "[$(Get-Date)] The file '$FileName' has been deleted successfully."
	Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
	Write-Host $LogLine -ForegroundColor "Green"

	$LogLine = "[$(Get-Date)] Process completed successfully. Waiting for new files..."
	Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
	Write-Host $LogLine -ForegroundColor "White"

	$Word.Quit()
}

# File System Watcher - Event Listener
Register-ObjectEvent $FileSystemWatcher "Created" -Action $Convert_WORD_to_PDF

$LogLine = "[$(Get-Date)] Word to PDF Converter 1.0.0 Started."
Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
Write-Host $LogLine -ForegroundColor "White"

$LogLine = "[$(Get-Date)] Listening for creation events in the 'Input' folder. Waiting for new files..."
Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
Write-Host $LogLine -ForegroundColor "White"

# Execution Loop
While ($True) {
	Start-Sleep 1
}

# Script Information
$Author = "Emil Kalantaryan"
$Name = "Word to PDF Converter"
$Version = "1.0.6"
$Date = "08/12/2021"

# Powershell Window Title
$Host.UI.RawUI.WindowTitle = "$Name - $Version"

# Script Information Output
Write-Host "--------------------------------------" -ForegroundColor "White"
Write-Host "  Author:   $Author" -ForegroundColor "White"
Write-Host "  Name:     $Name" -ForegroundColor "White"
Write-Host "  Version:  $Version" -ForegroundColor "White"
Write-Host "  Date:     $Date" -ForegroundColor "White"
Write-Host "--------------------------------------`n" -ForegroundColor "White"

# Checking if the required folders exist and creating them in case they do not exist
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

# File System Watcher - Declaration
$FileSystemWatcher = New-Object System.IO.FileSystemWatcher

# File System Watcher - Configuration
$FileSystemWatcher.Path = [System.Environment]::CurrentDirectory + "\Input"
$FileSystemWatcher.Filter = "*.doc*"
$FileSystemWatcher.IncludeSubdirectories = $False
$FileSystemWatcher.EnableRaisingEvents = $True
$FileSystemWatcher | Get-Member -MemberType Event

# Word to PDF Converter - Script Block
$Convert_Word_to_PDF = {
    # Word to PDF Converter - Configuration
    $InputPath = [System.Environment]::CurrentDirectory + "\Input"
    $FileTypes = "*.doc?"
    $Files = Get-ChildItem $InputPath -Filter $FileTypes
    $FilesCount = ($Files).Count

    # Checking if the 'Input' folder is empty or not
    if ("$Files" -ne "") {
        # Conversions Counter - Initialization with the value '0'
        $ConversionsCounter = 0

        # Word Application Object - Declaration
        $Word = New-Object -ComObject Word.Application

        # Word Application Object - Configuration
        $Word.Visible = $False
        $wdFormatPDF = 17

        if ($FilesCount -gt 1) {
            $LogLine = "`n[$(Get-Date)] New files detected in the 'Input' folder."
        } else {
            $LogLine = "`n[$(Get-Date)] New file detected in the 'Input' folder."
        }
        Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
        Write-Host $LogLine -ForegroundColor "Cyan"

        if ($FilesCount -gt 1) {
            $LogLine = "[$(Get-Date)] Starting the conversion process of the detected files.`n"
        } else {
            $LogLine = "[$(Get-Date)] Starting the conversion process of the detected file.`n"
        }
        Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
        Write-Host $LogLine -ForegroundColor "Cyan"

        # File conversion Loop
        Get-ChildItem $InputPath -Filter $FileTypes | ForEach-Object {
            # Current File information
            $FileName = "$_"
            $FilePath = $_.FullName

            # Converted file Output
            $OutputPath = [System.Environment]::CurrentDirectory + "\Output\" + $FileName.Substring(0,$FileName.LastIndexOf("."))

            $LogLine = "[$(Get-Date)] Generating PDF of the file: $FileName"
            Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
            Write-Host $LogLine -ForegroundColor "Yellow"

            # File conversion instructions
            $Document = $Word.Documents.Open($FilePath)
            $Document.SaveAs([ref] $OutputPath, [ref] $wdFormatPDF)
            $Document.Close()

            $LogLine = "[$(Get-Date)] The PDF of the file '$FileName' has been generated successfully."
            Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
            Write-Host $LogLine -ForegroundColor "Green"

            $LogLine = "[$(Get-Date)] Deleting the file: $FileName"
            Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
            Write-Host $LogLine -ForegroundColor "Yellow"

            # File delete instruction
            Remove-Item $FilePath

            $LogLine = "[$(Get-Date)] The file '$FileName' has been deleted successfully."
            Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
            Write-Host $LogLine -ForegroundColor "Green"

            # Conversions Counter - Update the value adding 1 to the previous value
            $ConversionsCounter++
        }

        if ($ConversionsCounter -gt 1) {
            $LogLine = "`n[$(Get-Date)] Operations in the queue completed."
        } else {
            $LogLine = "`n[$(Get-Date)] Operation in the queue completed."
        }
        Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
        Write-Host $LogLine -ForegroundColor "White"

        if ($ConversionsCounter -gt 1) {
            $LogLine = "[$(Get-Date)] $ConversionsCounter conversions of Word files to PDF have been performed."
        } else {
            $LogLine = "[$(Get-Date)] $ConversionsCounter conversion of Word file to PDF has been performed."
        }
        Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
        Write-Host $LogLine -ForegroundColor "White"

        if ($ConversionsCounter -gt 1) {
            $LogLine = "[$(Get-Date)] $ConversionsCounter Word files have been removed."
        } else {
            $LogLine = "[$(Get-Date)] $ConversionsCounter Word file has been removed."
        }
        Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
        Write-Host $LogLine -ForegroundColor "White"

        $LogLine = "[$(Get-Date)] Process completed successfully. Waiting for new files...`n"
        Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
        Write-Host $LogLine -ForegroundColor "White"

        # Word Application Close
        $Word.Quit()
    }
}

# File System Watcher - Event Listener
Register-ObjectEvent $FileSystemWatcher "Created" -Action $Convert_Word_to_PDF

$LogLine = "[$(Get-Date)] $Name $Version Started."
Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
Write-Host $LogLine -ForegroundColor "White"

$LogLine = "[$(Get-Date)] Listening for creation events in the 'Input' folder. Waiting for new files...`n"
Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
Write-Host $LogLine -ForegroundColor "White"

# Script Execution Loop
While ($True) {
	Start-Sleep 1
}

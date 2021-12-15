# Script Information
$Author = "Emil Kalantaryan"
$Name = "Word to PDF Converter"
$Version = "1.2.0"
$Date = "12/12/2021"

# Powershell Window Title
$Host.UI.RawUI.WindowTitle = "$Name $Version"

# Script Information Output
Write-Host "--------------------------------------`n" -ForegroundColor "White"
Write-Host "  Author:   $Author" -ForegroundColor "White"
Write-Host "  Name:     $Name" -ForegroundColor "White"
Write-Host "  Version:  $Version" -ForegroundColor "White"
Write-Host "  Date:     $Date`n" -ForegroundColor "White"
Write-Host "--------------------------------------`n" -ForegroundColor "White"

# Checking if the required folders & files exist and creating them in case they do not exist
if (!(Test-Path -Path ".\Input")) {
    $LogLine = "[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] Creating the folder: '$PSScriptRoot\Input'"
    Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
    Write-Host $LogLine -ForegroundColor "Cyan"
    New-Item -Path ".\Input" -ItemType "Directory" > $Null
}

if (!(Test-Path -Path ".\Output")) {
    $LogLine = "[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] Creating the folder: '$PSScriptRoot\Output'"
    Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
    Write-Host $LogLine -ForegroundColor "Cyan"
    New-Item -Path ".\Output" -ItemType "Directory" > $Null
}

if (!(Test-Path -Path ".\Word_to_PDF_Converter.json")) {
    $LogLine = "[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] Creating the file: '$PSScriptRoot\Word_to_PDF_Converter.json'"
    Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
    Write-Host $LogLine -ForegroundColor "Cyan"

    # A JSON Object with the required keys and values is created and stored in the file 'Word_to_PDF_Converter.json'
    $JSON = @{}
    $JSON.Add("CurrentSessionConversions", 0)
    $JSON.Add("TotalConversions", 0)
    $JSON | ConvertTo-Json | Out-File ".\Word_to_PDF_Converter.json"
} else {
    # Resets the value of 'CurrentSessionConversions' to 0 and saves it in the file 'Word_to_PDF_Converter.json'.
    $JSON = Get-Content -Path ".\Word_to_PDF_Converter.json" | ConvertFrom-Json
    $JSON.CurrentSessionConversions = 0
    $JSON | ConvertTo-Json | Out-File ".\Word_to_PDF_Converter.json"
}

# Powershell Window Title - Update
$Host.UI.RawUI.WindowTitle = "$Name $Version  -  Current session conversions: " + $JSON.CurrentSessionConversions + "  -  Total conversions: " + $JSON.TotalConversions

# File System Watcher - Declaration
$FileSystemWatcher = New-Object System.IO.FileSystemWatcher

# File System Watcher - Configuration
$FileSystemWatcher.Path = ".\Input"
$FileSystemWatcher.Filter = "*.doc*"
$FileSystemWatcher.IncludeSubdirectories = $False
$FileSystemWatcher.EnableRaisingEvents = $True
$FileSystemWatcher | Get-Member -MemberType Event

# Word to PDF Converter - Script Block
$Convert_Word_to_PDF = {
    # Getting files with extension '.doc' and '.docx' from the 'Input' folder
    $Files = Get-ChildItem ".\Input" -Filter "*.doc?" | Where-Object {$_.Extension -in @(".doc",".docx")}

    # Checking if the 'Input' folder contains files with extension '.doc' and '.docx' (ignoring Word temporary files)
    if (("$Files" -ne "") -and ($Event.SourceEventArgs.Name -notlike "~$*")) {
        # Conversions Counter - Setting initial value to '0'
        $ConversionsCounter = 0

        # Getting the number of files with '.doc' and '.docx' extension in the 'Input' folder
        $FilesCount = ($Files).Count

        # Reads the file 'Word_to_PDF_Converter.json' and collects the required keys and values to iterate over them
        $JSON = Get-Content -Path ".\Word_to_PDF_Converter.json" | ConvertFrom-Json

        # Word Application Object - Declaration
        $Word = New-Object -ComObject Word.Application

        # Word Application Object - Setting Microsoft Word Window visibility to 'False'
        $Word.Visible = $False

        if ($FilesCount -gt 1) {
            $LogLine = "`n[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] $FilesCount new files detected in the 'Input' folder."
        } else {
            $LogLine = "`n[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] $FilesCount new file detected in the 'Input' folder."
        }
        Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
        Write-Host $LogLine -ForegroundColor "Cyan"

        if ($FilesCount -gt 1) {
            $LogLine = "[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] Starting the conversion process of the $FilesCount detected files.`n"
        } else {
            $LogLine = "[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] Starting the conversion process of the detected file.`n"
        }
        Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
        Write-Host $LogLine -ForegroundColor "Cyan"

        # Word Files conversion Loop
        $Files | ForEach-Object {
            $LogLine = "[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] Generating PDF of the file: $_"
            Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
            Write-Host $LogLine -ForegroundColor "Yellow"

            # File conversion instructions
            $Document = $Word.Documents.Open($_.FullName)
            $Document.SaveAs([ref] $([System.Environment]::CurrentDirectory + "\Output\" + "$_".Substring(0,"$_".LastIndexOf("."))), [ref] 17) # 17 = wdFormatPDF - PDF format. (https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveformat)
            $Document.Close()

            $LogLine = "[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] The PDF of the file '$_' has been generated successfully."
            Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
            Write-Host $LogLine -ForegroundColor "Green"

            $LogLine = "[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] Deleting the file: $_"
            Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
            Write-Host $LogLine -ForegroundColor "Yellow"

            # File delete instruction
            Remove-Item $_.FullName

            $LogLine = "[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] The file '$_' has been deleted successfully."
            Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
            Write-Host $LogLine -ForegroundColor "Green"

            # Conversions Counter, Current Session Conversions and Total Conversions - Update the values adding 1 to the previous value
            $ConversionsCounter++
            $JSON.CurrentSessionConversions++
            $JSON.TotalConversions++
            
            # Powershell Window Title - Update
            $Host.UI.RawUI.WindowTitle = "$Name $Version  -  Current session conversions: " + $JSON.CurrentSessionConversions + "  -  Total conversions: " + $JSON.TotalConversions
        }
        # Saving in the file 'Word_to_PDF_Converter.json' the new values of 'CurrentSessionConversions' and 'TotalConversions'
        $JSON | ConvertTo-Json | Out-File ".\Word_to_PDF_Converter.json"

        if ($ConversionsCounter -gt 1) {
            $LogLine = "`n[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] Operations in the queue completed."
        } else {
            $LogLine = "`n[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] Operation in the queue completed."
        }
        Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
        Write-Host $LogLine -ForegroundColor "White"

        if ($ConversionsCounter -gt 1) {
            $LogLine = "[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] $ConversionsCounter conversions of Word files to PDF have been performed."
        } else {
            $LogLine = "[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] $ConversionsCounter conversion of Word file to PDF has been performed."
        }
        Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
        Write-Host $LogLine -ForegroundColor "White"

        if ($ConversionsCounter -gt 1) {
            $LogLine = "[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] $ConversionsCounter Word files have been removed."
        } else {
            $LogLine = "[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] $ConversionsCounter Word file has been removed."
        }
        Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
        Write-Host $LogLine -ForegroundColor "White"

        # Stats Information Output
        Write-Host "`n----------------------------------------" -ForegroundColor "White"
        Write-Host "                 Stats`n" -ForegroundColor "White"
        Write-Host "  Current session conversions: " $JSON.CurrentSessionConversions -ForegroundColor "White"
        Write-Host "  Total conversions: " $JSON.TotalConversions -ForegroundColor "White"
        Write-Host "`n----------------------------------------`n" -ForegroundColor "White"

        $LogLine = "[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] Process completed successfully. Waiting for new files...`n"
        Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
        Write-Host $LogLine -ForegroundColor "White"

        # Word Application Close
        $Word.Quit()

        # Clear Variables
        Clear-Variable $ConversionsCounter
        Clear-Variable $Document
        Clear-Variable $Files
        Clear-Variable $FilesCount
        Clear-Variable $JSON
        Clear-Variable $LogLine
        Clear-Variable $Word
    }
}

# File System Watcher - Event Listener
Register-ObjectEvent $FileSystemWatcher "Created" -Action $Convert_Word_to_PDF

$LogLine = "[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] $Name $Version Started."
Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
Write-Host $LogLine -ForegroundColor "White"

$LogLine = "[$(Get-Date -Format "dd/MM/yyy HH:mm:ss")] Listening for creation events in the 'Input' folder. Waiting for new files...`n"
Add-content ".\Word_to_PDF_Converter.log" -Value $LogLine
Write-Host $LogLine -ForegroundColor "White"

# Default conversion of the files that have been added to the 'Input' folder while 'Word to PDF Converter' was not running
$FirstFile = Get-ChildItem ".\Input" -Filter "*.doc?" | Where-Object {$_.Extension -in @(".doc",".docx")} | Select-Object -First 1

# Checking if the first file of the 'Input' folder has '.doc' and '.docx' extension (ignoring Word temporary files)
if (("$FirstFile" -ne "") -and ("$FirstFile" -notlike "~$*")) {
    # Moves the first file of the 'Input' folder to a temporary folder and returns it to the original location for activate the File System Event Listener
    New-Item -Path ".\Input\Temp" -ItemType "Directory"
    Move-Item -Path ".\Input\$FirstFile" -Destination ".\Input\Temp\$FirstFile"
    Move-Item -Path ".\Input\Temp\$FirstFile" -Destination ".\Input\$FirstFile"
    Remove-Item ".\Input\Temp"
}

# Removing variables that have already been used and will not be used again in the current runtime
Remove-Variable $Author
Remove-Variable $Date
Remove-Variable $FirstFile
Remove-Variable $JSON
Remove-Variable $LogLine

# Script Execution Loop
While ($True) {
    Start-Sleep -Milliseconds 500
}

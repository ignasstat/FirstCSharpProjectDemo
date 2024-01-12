function Check-FileExistsEFT {
    param(
        [long]$lngFileID,
        [string]$JobNumber,
        [bool]$HeaderInFirstRow,
        [string]$strFile
    )

    $FileExistsTriggerFolder = "YOUR_TRIGGER_FOLDER_PATH"
    $fileExistsTriggerToProcessFolder = Join-Path -Path $FileExistsTriggerFolder -ChildPath "ToProcess"
    $fileExistsTriggerFullPath = Join-Path -Path $fileExistsTriggerToProcessFolder -ChildPath "$lngFileID.trg"
    $foundFolder = Join-Path -Path $FileExistsTriggerFolder -ChildPath "Found"
    $foundNotEmpty = Join-Path -Path $foundFolder -ChildPath "NotEmpty"
    $foundEmpty = Join-Path -Path $foundFolder -ChildPath "Empty"
    $notFoundFolder = Join-Path -Path $FileExistsTriggerFolder -ChildPath "NotFound"
    
    $fileFoundNotEmptyPath = Join-Path -Path $foundNotEmpty -ChildPath "$lngFileID.trg"
    $fileFoundEmptyPath = Join-Path -Path $foundEmpty -ChildPath "$lngFileID.trg"
    $fileNotFoundPath = Join-Path -Path $notFoundFolder -ChildPath "$lngFileID.trg"

    # Delete existing trigger files
    Remove-Item -Path $fileFoundNotEmptyPath -ErrorAction SilentlyContinue
    Remove-Item -Path $fileFoundEmptyPath -ErrorAction SilentlyContinue
    Remove-Item -Path $fileNotFoundPath -ErrorAction SilentlyContinue

    # EFT 769 - Adding additional check if a file already exists.
    Remove-Item -Path $fileExistsTriggerFullPath -ErrorAction SilentlyContinue

    # Create a trigger file called by fileID
    $strHeaderInFirstRow = $HeaderInFirstRow.ToString()

    $content = @"
$strFile
$JobNumber
$strHeaderInFirstRow
"@

    $content | Out-File -FilePath $fileExistsTriggerFullPath -Force

    # Loop until trigger file appears in one of the three folders
    while (-not (Test-Path $fileFoundNotEmptyPath -or Test-Path $fileFoundEmptyPath -or Test-Path $fileNotFoundPath)) {
        # Do nothing, just wait for the trigger file
    }

    # Check if the trigger file was moved to notfound, FoundNotEmpty, or FoundEmpty subfolders
    if (Test-Path $fileNotFoundPath) {
        "not found"
        Remove-Item -Path $fileNotFoundPath -ErrorAction SilentlyContinue
    } elseif (Test-Path $fileFoundNotEmptyPath) {
        "OK"
        Remove-Item -Path $fileFoundNotEmptyPath -ErrorAction SilentlyContinue
    } elseif (Test-Path $fileFoundEmptyPath) {
        "Is empty"
        Remove-Item -Path $fileFoundEmptyPath -ErrorAction SilentlyContinue
    }
}

# Example usage:
$lngFileID = 123
$JobNumber = "YourJobNumber"
$HeaderInFirstRow = $true
$strFile = "C:\Path\To\Your\File.txt"

Check-FileExistsEFT -lngFileID $lngFileID -JobNumber $JobNumber -HeaderInFirstRow $HeaderInFirstRow -strFile $strFile

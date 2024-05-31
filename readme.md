function CheckFileExists_EFT {
    param (
        [int]$lngFileID,
        [string]$jobNumber,
        [bool]$headerInFirstRow,
        [string]$strFile
    )
    
    # Define base folders
    $fileExistsTriggerFolder = "PathToYourTriggerFolder"  # Update this path as necessary
    $toProcessFolder = Join-Path $fileExistsTriggerFolder "ToProcess"
    $foundFolder = Join-Path $fileExistsTriggerFolder "Found"
    $notFoundFolder = Join-Path $fileExistsTriggerFolder "NotFound"
    
    # Define specific file paths
    $triggerFilePath = Join-Path $toProcessFolder "$lngFileID.trg"
    $foundNotEmptyPath = Join-Path $foundFolder "NotEmpty\$lngFileID.trg"
    $foundEmptyPath = Join-Path $foundFolder "Empty\$lngFileID.trg"
    $fileNotFoundPath = Join-Path $notFoundFolder "$lngFileID.trg"
    
    # Delete existing trigger files
    Remove-Item -Path $foundNotEmptyPath, $foundEmptyPath, $fileNotFoundPath, $triggerFilePath -ErrorAction SilentlyContinue
    
    # Create a new trigger file with details
    [System.IO.File]::WriteAllLines($triggerFilePath, @($strFile, $jobNumber, $headerInFirstRow))
    
    # Check for trigger file appearance in specified folders
    $maxAttempts = 3
    for ($attemptCount = 0; $attemptCount -lt $maxAttempts; $attemptCount++) {
        Start-Sleep -Seconds 3
        
        if (Test-Path $foundNotEmptyPath -or Test-Path $foundEmptyPath -or Test-Path $fileNotFoundPath) {
            break
        }
    }

    # Handling file found scenarios
    if ($attemptCount -eq $maxAttempts) {
        Write-Host "Trigger file not found after multiple attempts."
        return "trigger notfound"
    }

    if (Test-Path $fileNotFoundPath) {
        Remove-Item $fileNotFoundPath -ErrorAction SilentlyContinue
        return "not found"
    } elseif (Test-Path $foundNotEmptyPath) {
        Remove-Item $foundNotEmptyPath -ErrorAction SilentlyContinue
        return "OK"
    } elseif (Test-Path $foundEmptyPath) {
        Remove-Item $foundEmptyPath -ErrorAction SilentlyContinue
        return "Is empty"
    }
}

# Example Usage:
$lngFileID = 12345
$jobNumber = 'Job001'
$headerInFirstRow = $true
$strFile = 'example.txt'
$result = CheckFileExists_EFT -lngFileID $lngFileID -jobNumber $jobNumber -headerInFirstRow $headerInFirstRow -strFile $strFile
Write-Host "Result: $result"

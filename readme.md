function CheckFileExists_EFT([int]$lngFileID, [string]$jobNumber, [bool]$headerInFirstRow, [string]$strFile) {
    $fileExistsTriggerFolder = "PathToYourTriggerFolder"  # Set this to your actual trigger folder path
    $fileExistsTriggerToProcessFolder = "$fileExistsTriggerFolder\ToProcess\"
    $fileExistsTriggerFullPath = "${fileExistsTriggerToProcessFolder}${lngFileID}.trg"
    
    $foundFolder = "$fileExistsTriggerFolder\Found\"
    $foundNotEmpty = "${foundFolder}NotEmpty\"
    $foundEmpty = "${foundFolder}Empty\"
    $notFoundFolder = "$fileExistsTriggerFolder\NotFound\"
    
    $fileFoundNotEmptyPath = "${foundNotEmpty}${lngFileID}.trg"
    $fileFoundEmptyPath = "${foundEmpty}${lngFileID}.trg"
    $fileNotFoundPath = "${notFoundFolder}${lngFileID}.trg"
    
    # Delete existing trigger files
    DeleteFile $fileFoundNotEmptyPath
    DeleteFile $fileFoundEmptyPath
    DeleteFile $fileNotFoundPath
    DeleteFile $fileExistsTriggerFullPath

    # Create a new trigger file
    [System.IO.File]::WriteAllLines($fileExistsTriggerFullPath, @($strFile, $jobNumber, $headerInFirstRow))
    
    # Check for trigger file appearance in folders
    $attemptCount = 0
    $maxAttempts = 3
    
    do {
        Start-Sleep -Seconds 3
        if ((Test-Path $fileFoundNotEmptyPath) -or (Test-Path $fileFoundEmptyPath) -or (Test-Path $fileNotFoundPath)) {
            break
        }
        $attemptCount++
    } while ($attemptCount -lt $maxAttempts)

    if ($attemptCount -eq $maxAttempts) {
        Write-Host "Trigger file not found after multiple attempts."
        return "trigger notfound"
    }

    if (Test-Path $fileNotFoundPath) {
        DeleteFile $fileNotFoundPath
        return "not found"
    } elseif (Test-Path $fileFoundNotEmptyPath) {
        DeleteFile $fileFoundNotEmptyPath
        return "OK"
    } elseif (Test-Path $fileFoundEmptyPath) {
        DeleteFile $fileFoundEmptyPath
        return "Is empty"
    }
}

function Remove-FileWithRetry {
    param (
        [string]$filePath,
        [int]$maxAttempts = 6,
        [int]$retryIntervalSeconds = 10
    )

    # Create a FileSystemObject-like instance using Get-Item (native PowerShell approach)
    if (-Not (Test-Path $filePath)) {
        Write-Host "File does not exist: $filePath"
        return
    }

    $attemptCount = 0
    $success = $false

    do {
        try {
            # Attempt to delete the file
            Remove-Item $filePath -ErrorAction Stop
            $success = $true
            Write-Host "File deleted successfully: $filePath"
            break
        } catch {
            # Log the error and wait before retrying
            Write-Host "Failed to delete file: $filePath. Error: $_"
            Start-Sleep -Seconds $retryIntervalSeconds
        }

        $attemptCount++
    } while ($attemptCount -lt $maxAttempts)

    # Check if deletion was not successful after all attempts
    if (-not $success) {
        Write-Host "Trigger file '$filePath' could not be deleted after multiple attempts."
    }
}

# Usage example:
$filePath = "C:\path\to\your\file.txt"
Remove-FileWithRetry -filePath $filePath

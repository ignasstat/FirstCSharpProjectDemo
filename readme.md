function Delete-File {
    param(
        [string]$filePath
    )

    $fso = New-Object -ComObject Scripting.FileSystemObject

    # Check if the file exists before attempting to delete
    if (-not $fso.FileExists($filePath)) {
        Write-Host "File '$filePath' does not exist."
        return
    }

    $attemptCount = 0
    $maxAttempts = 6
    $success = $false

    do {
        # Attempt to delete the file
        try {
            $fso.DeleteFile($filePath)
        } catch {
            # Ignore errors during deletion
        }

        # Check if the file still exists after deletion attempt
        if (-not $fso.FileExists($filePath)) {
            $success = $true
            break
        }

        # Wait for 10 seconds before the next deletion attempt
        Start-Sleep -Seconds 10

        $attemptCount++
    } while ($attemptCount -lt $maxAttempts)

    # Message if the file could not be deleted after multiple attempts
    if (-not $success) {
        Write-Host "Trigger file '$filePath' could not be deleted after multiple attempts."
    }
}

# Example usage:
$filePath = "C:\Path\To\Your\File.txt"
Delete-File -filePath $filePath

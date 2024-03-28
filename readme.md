function Check-FileExistsEFT {
    param(
        [int]$lngFileID,
        [string]$JobNumber,
        [bool]$HeaderInFirstRow,
        [string]$strFile
    )

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
    Remove-Item $fileFoundNotEmptyPath -ErrorAction SilentlyContinue
    Remove-Item $fileFoundEmptyPath -ErrorAction SilentlyContinue

    # Main code    
    $FileCheck = CheckFileExists -FileID $lngFileID -JobNumber $JobNumber

    # Assuming $blnProcess is determined earlier in the script.
    if ($blnProcess -eq $true) {
        # Database operations should be replaced with appropriate PowerShell commands or functions
        # This is a placeholder for how you might perform a database query in PowerShell
        $query = "SELECT * FROM dbo.vw_CT_CheckFile WHERE fileid = $lngFileID"
        $databaseResult = Invoke-SqlQuery -Query $query -Database "YourDatabase" -ServerInstance "YourServerInstance"

        if ($databaseResult) {
            $fileUser = $databaseResult.FileUser.Trim().ToLower()
            if ($fileUser -and $fileUser -ne $strUserName.Trim().ToLower()) {
                Write-Host "File already in use by $fileUser"
                $blnProcess = $false
            } elseif (-not $fileUser) {
                # Lock the file so nobody else can use it
                # This is a placeholder for inserting data into a database
                $insertQuery = @"
                INSERT INTO dbo.CT_FileLock (FileID, Source, FileUser, dts) 
                VALUES ($lngFileID, 'SourcePlaceholder', '$strUserName', GETDATE())
"@
                Invoke-SqlQuery -Query $insertQuery -Database "YourDatabase" -ServerInstance "YourServerInstance"
            }
        } else {
            Write-Host "File no longer available to process, reject it and refresh data."
            # Placeholder for refreshing data
        }

        if ($blnProcess) {
            # Further processing...
        }
    }
}


# Assuming $blnProcess is true and further processing is allowed
if ($blnProcess -eq $true) {
    $source = if ($databaseResult.Source -eq "E") { "E" } else { "I" }
    $createdDate = [DateTime]$databaseResult.CreatedDate
    $updatedDate = [DateTime]$databaseResult.UpdatedDate
    $strDate = if ($createdDate -gt $updatedDate) { $createdDate.ToString("dd MMM yyyy HH:mm") } else { $updatedDate.ToString("dd MMM yyyy HH:mm") }

    # Prepare the SQL for inserting a new job run entry
    $insertJobRunSql = @"
    INSERT INTO dbo.CT_JobRun (CT_JobID, fileid, FileSource, RunNo, CreatedBy, createddate, RunStatus, DueByDate)
    SELECT ct_JobID,
           $lngFileID,
           '$source',
           CASE WHEN LastRun > LastNeptuneRun THEN LastRun + 1 ELSE LastNeptuneRun + 1 END,
           '$strUserName',
           GETDATE(),
           'Logged',
           dbo.fn_CallTraceDueDate_New('$strDate', j.job_number)
    FROM dbo.vw_CallTraceJobList v
    INNER JOIN dbo.ct_jobs j ON v.job_number = j.job_number
    WHERE v.job_number = '$JobNumber'
"@

    # Execute the SQL command to insert the job run entry
    Invoke-SqlQuery -Query $insertJobRunSql -Database "YourDatabase" -ServerInstance "YourServerInstance"

    Write-Host "Linking file ID $lngFileID to job number $JobNumber"
}

# Note: This assumes the presence of a function or cmdlet named `Invoke-SqlQuery` which is a placeholder.
# You will need to replace it with actual code or a function that executes SQL queries against your database.

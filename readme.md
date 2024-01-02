# Variables
$Source = ""
$strDate = ""
$CreatedDate = ""
$UpdatedDate = ""
$FileSize = 0
$blnProcess = $false
$strSQL = ""
$rst = New-Object -ComObject ADODB.Recordset
$selectedRow = 0
$JobNumber = ""
$IsEmpty = $false
$FileCheck = ""

# Function to check file existence
function CheckFileExists($lngFileID, $JobNumber) {
    # Implementation of the CheckFileExists function in PowerShell
    # ...

    return "OK"  # Replace with actual implementation
}

# Function to clear file lock
function ClearFileLock($FileID, $Source) {
    # Implementation of the ClearFileLock function in PowerShell
    # ...
}

# Main code
if (-not [System.Management.Automation.LanguagePrimitives]::IsNull($lstFileView.Value)) {
    # Job number from list
    $lngFileID = $lstFileView.Value
    $blnProcess = $true

    # Job number from list
    $selectedRow = $lstFileView.ListIndex
    $selectedRow++
    $JobNumber = $lstFileView.Column(1, $selectedRow)

    if ($JobNumber -eq "No Matches" -or $JobNumber -eq "Multiple Matches") {
        $blnProcess = $false
        Write-Host "File has $JobNumber with jobs"
    }

    # CTC116 Additional parameter JobNumber, also function returns string instead of boolean
    $FileCheck = CheckFileExists $lngFileID $JobNumber

    if ($blnProcess -eq $true) {
        # Check if the file has been processed or selected by another user
        $strSQL = "select * from dbo.vw_CT_CheckFile where fileid = $($lngFileID)"
        $rst.Open($strSQL, $db, 1, 1)

        if (-not $rst.EOF) {
            if (-not [System.Management.Automation.LanguagePrimitives]::IsNull($rst.Fields("FileUser").Value)) {
                if ($rst.Fields("FileUser").Value.Trim().ToLower() -eq $strUserName.Trim().ToLower()) {
                    # Current user, so OK
                } else {
                    Write-Host "File already in use by $($rst.Fields("FileUser").Value)"
                    $blnProcess = $false
                }
            } else {
                # Lock the file so nobody else can use it
                $strSQL = "insert into dbo.CT_FileLock (FileID, Source, FileUser, dts) values ($($rst.Fields("FileID").Value), '$($rst.Fields("Source").Value)', '$strUserName', getdate())"
                $db.Execute($strSQL)
            }
        } else {
            $blnProcess = $false
            Write-Host "File no longer available to process, Reject it and Refresh Data"
            $lstFileView.Requery()
            $lstRejected.Requery()
        }

        # CTC 039 - Cannot check if the file exists directly anymore as the user won't have permission to the EFT folder, added method CheckFileExists
        if ($blnProcess -eq $true) {
            # Need to check if the file has changed since the last update of the list.
            # If SFTP
            if ($rst.Fields("Source").Value -eq "E") {
                if (-not $FileCheck -eq "OK") {
                    Write-Host "This file or version of it $FileCheck, Reject it and Refresh Data"
                    ClearFileLock $rst.Fields("FileID").Value "E"
                    $blnProcess = $false
                }
            } else {
                # If Data In
                if (-not $FileCheck -eq "OK") {
                    Write-Host "This file or version of it $FileCheck, Reject it and Refresh Data"
                    ClearFileLock $rst.Fields("FileID").Value "I"
                    $blnProcess = $false
                }
            }
        }

        if ($blnProcess -eq $true) {
            if ($rst.Fields("Source").Value -eq "E") {
                $Source = "E"
            } else {
                $Source = "I"
            }

            $CreatedDate = $rst.Fields("CreatedDate").Value
            $UpdatedDate = $rst.Fields("UpdatedDate").Value

            if ($CreatedDate -gt $UpdatedDate) {
                $strDate = $CreatedDate.ToString("dd MMM yyyy HH:mm")
            } else {
                $strDate = $UpdatedDate.ToString("dd MMM yyyy HH:mm")
            }

            # Create an entry in the Job Run Table
            $strSQL = "insert into dbo.CT_JobRun (CT_JobID,fileid,FileSource,RunNo,CreatedBy,createddate,RunStatus, DueByDate) "
            $strSQL += "Select ct_JobID, "
            $strSQL += "$($lngFileID),"
            $strSQL += "'$Source',"
            $strSQL += "case when LastRun > LastNeptuneRun then LastRun+1 else LastNeptuneRun+1 end, "
            $strSQL += "'$strUserName',"
            $strSQL += "getdate(),"
            $strSQL += "'Logged',"
            $strSQL += "dbo.fn_CallTraceDueDate_New('$strDate', j.job_number ) "
            $strSQL += "from dbo.vw_CallTraceJobList v "
            $strSQL += "inner join dbo.ct_jobs j on v.job_number = j.job_number "
            $strSQL += "where v.job_number = '$JobNumber'"

            Write-Host "Linking $($strFileName) to $($JobNumber)"
            $db.Execute($strSQL)
        }

        $rst.Close()
    }
}

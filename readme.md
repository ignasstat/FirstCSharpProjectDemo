function Check-FileExists {
    param(
        [int]$lngFileID,
        [string]$JobNumber
    )

    $dbConnectionString = "YOUR_DATABASE_CONNECTION_STRING"
    $FileExistsTriggerFolder = "YOUR_TRIGGER_FOLDER_PATH"

    $connection = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $dbConnectionString
    $connection.Open()

    $command = $connection.CreateCommand()
    $command.CommandText = "select * from dbo.vw_ExistingFiles where fileid = $lngFileID"
    $reader = $command.ExecuteReader()

    if ($reader.HasRows) {
        $reader.Read()
        $strFile = $reader["Folder"] + $reader["Filename"]
        $reader.Close()

        $HeaderInFirstRow = Get-HeaderInFirstRow -JobNumber $JobNumber

        if ($strFile -like "\\cig.local\Data\Marketing Solutions Departments\Production\*") {
            $result = Check-FileExistsLocal -myFile $strFile -HeaderInFirstRow $HeaderInFirstRow
        } else {
            $result = Check-FileExistsEFT -lngFileID $lngFileID -JobNumber $JobNumber -HeaderInFirstRow $HeaderInFirstRow -strFile $strFile
        }

        $result
    } else {
        $reader.Close()
        Write-Host "file no longer in view (someone else may have processed it)"
        "not in view"
    }

    $connection.Close()
}

function Get-HeaderInFirstRow {
    param(
        [string]$JobNumber
    )

    $dbConnectionString = "YOUR_DATABASE_CONNECTION_STRING"

    $connection = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $dbConnectionString
    $connection.Open()

    $command = $connection.CreateCommand()
    $command.CommandText = @"
    SELECT HeadersInFirstRow
    FROM neptunefileimporter.[fileimporter].[Jobs]
    WHERE JobId IN (SELECT Max(JobId) FROM neptunefileimporter.[fileimporter].[Jobs] WHERE DestinationTable = '$JobNumber')
"@
    $reader = $command.ExecuteReader()

    if ($reader.HasRows) {
        $reader.Read()
        $headerInFirstRow = $reader["HeadersInFirstRow"]
        $reader.Close()
        $headerInFirstRow
    } else {
        $reader.Close()
        Write-Host "Error retrieving HeadersInFirstRow value from neptunefileimporter.[fileimporter].[Jobs]"
        $false
    }

    $connection.Close()
}

function Check-FileExistsLocal {
    param(
        [string]$myFile,
        [bool]$HeaderInFirstRow
    )

    $fso = New-Object -ComObject Scripting.FileSystemObject

    if (-not $fso.FileExists($myFile)) {
        "not found"
    } else {
        $mf = $fso.GetFile($myFile)

        if ($mf.Size -gt 0) {
            if ($HeaderInFirstRow) {
                $result = Check-ContainsMultipleLines -FileName $myFile
                if ($result) {
                    "OK"
                } else {
                    "Is empty"
                }
            } else {
                "OK"
            }
        } else {
            "Is empty"
        }
    }
}

function Check-ContainsMultipleLines {
    param(
        [string]$FileName
    )

    $fso = New-Object -ComObject Scripting.FileSystemObject
    $mf = $fso.OpenTextFile($FileName, 1)

    $lineCount = 0

    while (-not $mf.AtEndOfStream -and $lineCount -lt 2) {
        $mf.ReadLine()
        $lineCount++
    }

    $mf.Close()

    if ($lineCount -gt 1) {
        $true
    } else {
        $false
    }
}

function Check-FileExistsEFT {
    param(
        [int]$lngFileID,
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
    Remove-Item -Path $fileFoundEmptyPath -

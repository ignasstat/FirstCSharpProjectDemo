$Global:dbconnectionString = "Data Source=pllwinlvsql002\mb21,1433;Integrated Security=SSPI;Initial Catalog=DataBureauDataLoadAudit"

function Get-SQLScalar {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [String] $Query
    )

    $Connection = New-Object System.Data.SQLClient.SQLConnection($dbconnectionString)
    $Connection.Open()
    $Command = New-Object System.Data.SQLClient.SQLCommand($Query, $Connection)
    $SQLScalar = $Command.ExecuteScalar()
    $Connection.Close()

    return $SQLScalar
}

function Get-HeaderInFirstRow {
    param (
        [string]$jobNumber
    )

    $sql = "SELECT HeadersInFirstRow FROM neptunefileimporter.[fileimporter].[Jobs] WHERE JobId IN (SELECT Max(JobId) FROM neptunefileimporter.[fileimporter].[Jobs] WHERE DestinationTable = '$jobNumber')"
    Write-Host $sql
    $result = Get-SQLScalar -Query $sql 
    Write-Host "Result: $result"

    if ($result -ne $null) {
        return $result
    } else {
        Write-Host "Error retrieving HeadersInFirstRow value from fileimporter.Jobs"
        return $false
    }
}

function CheckFileExists_Local {
    param (
        [string]$myFile,
        [bool]$headerInFirstRow
    )

    if (-Not (Test-Path $myFile)) {
        return "not found"
    } else {
        $fileInfo = Get-Item $myFile
        if ($fileInfo.Length -gt 0) {
            if ($headerInFirstRow) {
                if (ContainsMultipleLines $myFile) {
                    return "OK"
                } else {
                    return "Is empty"
                }
            } else {
                return "OK"
            }
        } else {
            return "Is empty"
        }
    }
}

function ContainsMultipleLines {
    param (
        [string]$fileName
    )

    $lineCount = 0
    try {
        $reader = [System.IO.File]::OpenText($fileName)
        while (-not $reader.EndOfStream -and $lineCount -lt 2) {
            $reader.ReadLine()
            $lineCount++
        }
        $reader.Close()
    } catch {
        Write-Host "Error reading file $fileName"
    }
    return $lineCount -gt 1
}

$jobNumber = 'CDA100966'
$headerInFirstRow = Get-HeaderInFirstRow -jobNumber $jobNumber

$file = "\\cig.local\data\AppData\SFTP\Data\Usr\DataBureau\Configuration\Scripts\Test\CallTrace Console\CTC124_125\ToProcess\123.trg"
$ats = CheckFileExists_Local -myFile $file -headerInFirstRow $headerInFirstRow

Write-Host "File check result: $ats"

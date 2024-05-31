$Global:dbconnectionString = "Data Source=pllwinlvsql002\mb21,1433;Integrated Security=SSPI;Initial Catalog=DataBureauDataLoadAudit"
function Get-SQLScalar {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)] [String] $Query
    )

    $Connection = New-Object System.Data.SQLClient.SQLConnection($dbconnectionString)
    $Connection.Open()
    $Command = New-Object System.Data.SQLClient.SQLCommand($Query, $Connection)
    $SQLScalar = $Command.ExecuteScalar()
    $Connection.Close()

    return $SQLScalar
    }

Function Get-HeaderInFirstRow {
    param (
        [string]$jobNumber
    )

    $sql = "SELECT HeadersInFirstRow FROM neptunefileimporter.[fileimporter].[Jobs] WHERE JobId IN (SELECT Max(JobId) FROM neptunefileimporter.[fileimporter].[Jobs] WHERE DestinationTable = '$jobNumber'"
    Write-Host $sql
    $result = Get-SQLScalar -Query $sql 
    write-host $result.Rows.Count

    if ($result.Rows.Count -gt 0) {
        return $result
    } else {
        Write-Host "Error retrieving HeadersInFirstRow value from fileimporter.Jobs"
        return $false
    }
}

function CheckFileExists_Local($myFile, $headerInFirstRow) {
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

function ContainsMultipleLines($fileName) {
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
$Heada = Get-HeaderInFirstRow -jobNumber $jobNumber

$file = "\\cig.local\data\AppData\SFTP\Data\Usr\DataBureau\Configuration\Scripts\Test\CallTrace Console\CTC124_125\ToProcess\123.trg"

$ats = CheckFileExists_Local -myFile $file -headerInFirstRow $Heada




SELECT HeadersInFirstRow FROM neptunefileimporter.[fileimporter].[Jobs] WHERE JobId IN (SELECT Max(JobId) FROM neptunefileimporter.[fileimporter].[Jobs] WHERE DestinationTable = 'CDA100966'
Exception calling "ExecuteScalar" with "0" argument(s): "Incorrect syntax near 'CDA100966'."
At line:12 char:5
+     $SQLScalar = $Command.ExecuteScalar()
+     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : NotSpecified: (:) [], MethodInvocationException
    + FullyQualifiedErrorId : SqlException
 
0
Error retrieving HeadersInFirstRow value from fileimporter.Jobs

param
(
    [String]$TriggerFile
)

$TriggerFile = "test.trg"

function Get-SQLScalar {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [String] $Query,

        [Parameter(Mandatory = $true)]
        [String] $ConnectionString
    )

    $Connection = New-Object System.Data.SQLClient.SQLConnection($ConnectionString)
    $Connection.Open()
    $Command = New-Object System.Data.SQLClient.SQLCommand($Query, $Connection)
    $SQLScalar = $Command.ExecuteScalar()
    $Connection.Close()

    return $SQLScalar
}

$TriggerFolder = "Test"
$FoundFolder = "Test"
$FoundRecords = Join-Path $FoundFolder "Records"
$FoundNoRecords = Join-Path $FoundFolder "NoRecords"
$FoundEmpty = Join-Path $FoundFolder "Empty"
$NotFoundFolder = "Test"

$TriggerPath = Join-Path $TriggerFolder $TriggerFile

$Parameters = Get-Content -LiteralPath $TriggerPath
$FilePath = $Parameters[0]
$JobNumber = $Parameters[1]

# If header expected, then counting records
$connectionString = "Test"
$query = @"
    SELECT HeadersInFirstRow
    FROM neptunefileimporter.[fileimporter].[Jobs]
    WHERE JobId IN (SELECT Max(JobId) FROM neptunefileimporter.[fileimporter].[Jobs] WHERE DestinationTable = '$JobNumber')
"@
$header = Get-SQLScalar -Query $query -ConnectionString $connectionString

# If a file is present
if (Test-Path $FilePath) {
    # If not empty
    $fileInfo = Get-Item $FilePath
    $filesize = $fileInfo.Length

    if ($filesize -gt 0) {
        # If it should have a header, then counting records if not, then moving directly to Records folder
        if ($header -eq 1) {
            $fileContent = Get-Content -Path $FilePath
            $recordCount = $fileContent.Count

            if ($recordCount > 1) {
                Copy-Item -LiteralPath $TriggerPath -Destination $FoundRecords
            } else {
                Copy-Item -LiteralPath $TriggerPath -Destination $FoundNoRecords
            }
        } else {
            Copy-Item -LiteralPath $TriggerPath -Destination $FoundRecords
        }
    } else {
        Copy-Item -LiteralPath $TriggerPath -Destination $FoundEmpty
    }
} else {
    # File not found
    Copy-Item -LiteralPath $TriggerPath -Destination $NotFoundFolder
}

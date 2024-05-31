function QuerySQL {
    Param(  [String] $sql,
            [String] $dbconnectionString
          )
    $connection = New-Object System.Data.SqlClient.SqlConnection($dbconnectionString)
    $command = $connection.CreateCommand()
    $command.CommandText = $sql
    $connection.Open()
    $result = $command.ExecuteReader()
    $table = New-Object System.Data.DataTable
    $table.Load($result)
    $connection.Close()

    return $table

}
$dbconnectionString = "Test"
$sql = "SELECT Folder, Filename FROM dbo.vw_ExistingFiles WHERE fileid = 348054"

$result = QuerySQL -sql $sql -dbconnectionString $dbconnectionString

foreach ($row in $result.Rows) {
    Write-Host "Folder: $($row['Folder']) - Filename: $($row['Filename'])"
}

if ($result.Rows.Count -gt 0) {
    $firstRow = $result.Rows[0]
    Write-Host "First row - Folder: $($firstRow['Folder']), Filename: $($firstRow['Filename'])"
}

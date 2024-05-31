function QuerySQL {
    Param(
        [String] $sql,
        [String] $dbconnectionString
    )
    try {
        $connection = New-Object System.Data.SqlClient.SqlConnection($dbconnectionString)
        $command = $connection.CreateCommand()
        $command.CommandText = $sql
        $connection.Open()
        $result = $command.ExecuteReader()
        $table = New-Object System.Data.DataTable
        $table.Load($result)
        $connection.Close()
        Write-Host "Rows loaded: $($table.Rows.Count)"
        return $table
    } catch {
        Write-Host "Error: $_"
        return $null
    }
}

$dbconnectionString = "Data Source=server_name;Integrated Security=SSPI;Initial Catalog=database_name"
$sql = "SELECT Folder, Filename FROM dbo.vw_ExistingFiles WHERE fileid = 348054"

$result = QuerySQL -sql $sql -dbconnectionString $dbconnectionString

if ($result -ne $null) {
    Write-Host "Number of rows in the result: $($result.Rows.Count)"
    foreach ($row in $result.Rows) {
        Write-Host "Folder: $($row['Folder']), Filename: $($row['Filename'])"
    }
} else {
    Write-Host "No data returned from the query."
}

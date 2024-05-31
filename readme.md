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

# Access data from the result
if ($result -ne $null -and $result.Rows.Count -gt 0) {
    foreach ($row in $result.Rows) {
        Write-Host "Folder: $($row['Folder']) - Filename: $($row['Filename'])"
    }

    # Storing specific data in a variable
    $firstFolder = $result.Rows[0]["Folder"]
    Write-Host "First folder retrieved: $firstFolder"
} else {
    Write-Host "No rows returned or an error occurred."
}

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
        Write-Host "Rows loaded inside function: $($table.Rows.Count)"
        
        # Print columns information
        foreach ($col in $table.Columns) {
            Write-Host "Column: $($col.ColumnName), Type: $($col.DataType)"
        }
        
        return $table
    } catch {
        Write-Host "Error: $_"
        return $null
    }
}

$dbconnectionString = "Test"
$sql = "SELECT Folder, Filename FROM dbo.vw_ExistingFiles WHERE fileid = 348054"

$result = QuerySQL -sql $sql -dbconnectionString $dbconnectionString

if ($result -ne $null) {
    Write-Host "Number of rows in the result: $($result.Rows.Count)"
    
    # Print rows information
    foreach ($row in $result.Rows) {
        Write-Host "Folder: $($row['Folder']), Filename: $($row['Filename'])"
    }
} else {
    Write-Host "No data returned from the query."
}

Rows loaded inside function: 1
Column: Folder, Type: string
Column: Filename, Type: string
Number of rows in the result: 0

function QuerySQL {
    Param(  [String] $sql,
            [String] $dbconnectionString
          )
    try{
        $connection = New-Object System.Data.SqlClient.SqlConnection($dbconnectionString)
        $command = $connection.CreateCommand()
        $command.CommandText = $sql
        $connection.Open()
        $result = $command.ExecuteReader()
        $table = New-Object System.Data.DataTable
        $table.Load($result)
        $connection.Close()
        Write-Host "Rows loaded $($table.Rows.Count)"
        foreach ($row in $result.Rows) {
            Write-Host "Folder: $($row['Folder']) - Filename: $($row['Filename'])"
        }

        return $table
        }
    catch {
    Write-Host "Error:$_"
    return $null
    }

}
$dbconnectionString = "Data Source=pllwinlvsql002\mb21,1433;Integrated Security=SSPI;Initial Catalog=DataBureauDataLoadAudit"
#$sql = "SELECT Folder, Filename FROM dbo.vw_ExistingFiles WHERE fileid = 348054"
$sql = "SELECT Folder, Filename FROM dbo.vw_ExistingFiles WHERE fileid = 348054"

$result = QuerySQL -sql $sql -dbconnectionString $dbconnectionString

Write-Host "Rows loaded $($table.Rows.Count)"

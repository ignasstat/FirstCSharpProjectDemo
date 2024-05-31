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


Write-Host $result

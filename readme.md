function Get-HeadersInFirstRow {
    param(
        [string]$JobNumber,
        [string]$dbConnectionString = "Data Source=pllwinlvsql002\mb21,1433;Integrated Security=SSPI;Initial Catalog=DataBureauDataLoadAudit"
    )

    # Establish the database connection
    $connection = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $dbConnectionString
    $connection.Open()

    try {
        # Create and configure the command
        $command = $connection.CreateCommand()
        $command.CommandText = @"
        SELECT HeadersInFirstRow
        FROM neptunefileimporter.[fileimporter].[Jobs]
        WHERE JobId IN (
            SELECT Max(JobId)
            FROM neptunefileimporter.[fileimporter].[Jobs]
            WHERE DestinationTable = '$JobNumber'
        )
"@
        # Execute the command and process results
        $reader = $command.ExecuteReader()

        if ($reader.HasRows) {
            $reader.Read()
            $headerInFirstRow = $reader["HeadersInFirstRow"]
            $reader.Close()

            return $headerInFirstRow
        } else {
            $reader.Close()
            Write-Host "Error retrieving HeadersInFirstRow value from neptunefileimporter.[fileimporter].[Jobs]"
            return $false
        }
    } catch {
        Write-Host "An error occurred: $_"
    } finally {
        $connection.Close()
    }
}

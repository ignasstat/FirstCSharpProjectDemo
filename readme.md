# Variables
$Source = ""
$strDate = ""
$CreatedDate = ""
$UpdatedDate = ""
$FileSize = 0
$blnProcess = $false
$strSQL = ""
$rst = New-Object -ComObject ADODB.Recordset
$selectedRow = 0
$JobNumber = ""
$IsEmpty = $false
$FileCheck = ""

function Get-FileViewData {
    $dbConnectionString = "Data Source=pllwinlvsql002\mb21,1433;Integrated Security=SSPI;Initial Catalog=DataBureauDataLoadAudit"

    $connection = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $dbConnectionString
    $connection.Open()

    $command = $connection.CreateCommand()
    $command.CommandText = "SELECT TOP 1 Source, CreatedDate, UpdatedDate FROM vw.CallTraceFilesToDisplay_New_Test"
    $reader = $command.ExecuteReader()

    if ($reader.HasRows) {
        $reader.Read()

        $global:Source = $reader["Source"]
        $global:CreatedDate = $reader["CreatedDate"]
        $global:UpdatedDate = $reader["UpdatedDate"]

        $reader.Close()
    } else {
        $reader.Close()
        Write-Host "No records found in the view."
    }

    $connection.Close()
}

# Example Usage:
Get-FileViewData
Write-Host "Source: $Source"
Write-Host "CreatedDate: $CreatedDate"
Write-Host "UpdatedDate: $UpdatedDate"

# Continue with the rest of your script...

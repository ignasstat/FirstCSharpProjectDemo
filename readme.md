Function sendEmail(strTo As String, strSubject As String, strTextBody As String, strCC As String, Optional attachments As Variant)

    Dim objMessage As Object
    Dim objConfig As Object
    Dim objFields As Object
    Dim i As Integer
    
    On Error Resume Next ' Turn on error handling

    Set objMessage = CreateObject("CDO.Message")

    ' Set configuration
    Set objConfig = CreateObject("CDO.Configuration")
    Set objFields = objConfig.Fields

    ' Set SMTP server details
    With objFields
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtprelay.cig.local"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 ' SMTP port
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' Send using SMTP
        .Update
    End With

    ' Configure message properties
    With objMessage
        Set .Configuration = objConfig
        .To = strTo
        .CC = strCC
        If EmailAs <> "" Then
            .From = EmailAs
        End If
        .Subject = strSubject
        .HTMLBody = strTextBody
        
        ' Attach files from the provided array (if it's an array)
        If IsArray(attachments) Then
            For i = LBound(attachments) To UBound(attachments)
                If Len(attachments(i)) > 0 Then
                    .AddAttachment attachments(i)
                End If
            Next i
        End If

        .Send
    End With

    ' Error handling
    If Err.Number <> 0 Then
        ' Handle the error here (you can display a message or log it)
        Debug.Print "Error sending email: " & Err.Description
    End If

    On Error GoTo 0 ' Reset error handling
    
    ' Clean up
    Set objMessage = Nothing
    Set objConfig = Nothing
    Set objFields = Nothing

End Function









function Compare-ExcelWithTxt {
    param(
        [string]$ExcelFilePath,
        [string]$TxtFilePath
    )

    # Load the Excel interop assembly
    Add-Type -AssemblyName Microsoft.Office.Interop.Excel

    # Create an Excel application object
    $excel = New-Object -ComObject Excel.Application

    # Open the Excel workbook
    $workbook = $excel.Workbooks.Open($ExcelFilePath)

    # Get the first worksheet
    $worksheet = $workbook.Worksheets.Item(1)

    # Read the Excel data into an array
    $excelData = @()
    $row = 2 # Assuming data starts from the second row
    while ($true) {
        $clientName = $worksheet.Cells.Item($row, 1).Value2
        if ([string]::IsNullOrEmpty($clientName)) {
            break
        }

        $actualValue = $worksheet.Cells.Item($row, 2).Value2
        $excelData += [PSCustomObject]@{
            ClientName = $clientName
            ActualValue = $actualValue
        }

        $row++
    }

    # Close the Excel workbook and application
    $workbook.Close()
    $excel.Quit()

    # Read the text file
    $txtData = Get-Content -Path $TxtFilePath | ForEach-Object {
        $clientName, $actualValue = $_ -split '\|'
        [PSCustomObject]@{
            ClientName = $clientName.Trim()
            ActualValue = $actualValue.Trim()
        }
    }

    # Compare Excel data with text file data
    $differences = Compare-Object -ReferenceObject $excelData -DifferenceObject $txtData -Property ClientName, ActualValue

    # Output the differences
    if ($differences.Count -eq 0) {
        Write-Host "No differences found."
    } else {
        Write-Host "Differences found:"
        $differences | Format-Table -AutoSize
    }
}

# Usage example:
# Compare-ExcelWithTxt -ExcelFilePath "C:\path\to\your\excel.xlsx" -TxtFilePath "C:\path\to\your\textfile.txt"






-----------------------------

# Create an instance of the Excel application object
$excel = New-Object -ComObject Excel.Application

# Set the visibility of the Excel application to false
$excel.Visible = $false

# Open the Excel file
$workbook = $excel.Workbooks.Open("\\valdotcsclu06s06\backoffice_shares$\dataoperations\eftcd\is_transfer\tot 523\callmonitoraudit.xlsx")

# Get the first worksheet in the workbook
$worksheet = $workbook.Sheets.Item(1)

# Get the used range of cells in the worksheet
$usedRange = $worksheet.UsedRange

# Loop through the rows in the used range
for ($row = 2; $row -le $usedRange.Rows.Count; $row++) {
    # Get the values in columns C, D, and E
    $columnCValue = $usedRange.Cells.Item($row, 3).Value2
    $columnDValue = $usedRange.Cells.Item($row, 4).Value2
    $columnEValue = $usedRange.Cells.Item($row, 5).Value2

    # Check if columns C and D have "N" values and if column E has "Y" value
    #Alternatively can use Switch
    if ($columnCValue -eq "N") {
        # Found a match
        Write-Host "Found a row with 'N' values in columns C and D, and 'Y' value in column E."
        Write-Host "Row: $row"
    }
     if ($columnDValue -eq "N") {
        # Found a match
        Write-Host "Found a row with 'N' values in columns C and D, and 'Y' value in column E."
        Write-Host "Row: $row"
    }
     if ($columnEValue -eq "Y") {
        # Found a match
        Write-Host "Found a row with 'N' values in columns C and D, and 'Y' value in column E."
        Write-Host "Row: $row"
    }
}

# Close the Excel file
$workbook.Close()

# Quit the Excel application
$excel.Quit()

# Release the Excel interop objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null








function Compare-ExcelWithTxt {
    param(
        [string]$ExcelFilePath,
        [string]$TxtFilePath
    )

    # Load the Excel interop assembly
    Add-Type -AssemblyName Microsoft.Office.Interop.Excel

    # Create an Excel application object
    $excel = New-Object -ComObject Excel.Application

    # Open the Excel workbook
    $workbook = $excel.Workbooks.Open($ExcelFilePath)

    # Get the first worksheet
    $worksheet = $workbook.Sheets.Item(1)

    # Get the used range of cells in the worksheet
    $usedRange = $worksheet.UsedRange

    # Initialize an array to store Excel data
    $excelData = @()

    # Loop through the rows in the used range
    for ($row = 2; $row -le $usedRange.Rows.Count; $row++) {
        $columnCValue = $usedRange.Cells.Item($row, 3).Value2
        $columnDValue = $usedRange.Cells.Item($row, 4).Value2
        $columnEValue = $usedRange.Cells.Item($row, 5).Value2

        $excelData += [PSCustomObject]@{
            ColumnC = $columnCValue
            ColumnD = $columnDValue
            ColumnE = $columnEValue
        }
    }

    # Close the Excel workbook and application
    $workbook.Close()
    $excel.Quit()

    # Release the Excel interop objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    # Read the text file
    $txtData = Get-Content -Path $TxtFilePath | ForEach-Object {
        $clientName, $actualValue = $_ -split '\|'
        [PSCustomObject]@{
            ClientName = $clientName.Trim()
            ActualValue = $actualValue.Trim()
        }
    }

    # Compare Excel data with text file data
    $differences = Compare-Object -ReferenceObject $excelData -DifferenceObject $txtData -Property ClientName, ActualValue

    # Output the differences
    if ($differences.Count -eq 0) {
        Write-Host "No differences found."
    } else {
        Write-Host "Differences found:"
        $differences | Format-Table -AutoSize
    }
}

# Usage example:
# Compare-ExcelWithTxt -ExcelFilePath "C:\path\to\your\excel.xlsx" -TxtFilePath "C:\path\to\your\textfile.txt"



# Replace 'YourFile.csv' with the path to your CSV or text file
$filePath = 'YourFile.csv'

# Use the Get-Content cmdlet to read the file into an array of lines
$fileContent = Get-Content -Path $filePath

# Use the Count property to get the number of records
$recordCount = $fileContent.Count

# Output the count
Write-Host "Number of records in $filePath: $recordCount"

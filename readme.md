foreach ($row in $result.Rows) {
    Write-Host "Folder: $($row['Folder']) - Filename: $($row['Filename'])"
}
if ($result.Rows.Count -gt 0) {
    $firstRow = $result.Rows[0]
    Write-Host "First row - Folder: $($firstRow['Folder']), Filename: $($firstRow['Filename'])"
}
$result | Export-Csv -Path "output.csv" -NoTypeInformation

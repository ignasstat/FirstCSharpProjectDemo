$HeaderInFirstRow = $true
$myFile = "\\cig.local\data\AppData\SFTP\Data\Usr\DataBureau\Configuration\Scripts\Test\CallTrace Console\test.txt"



function Check-FileExistsLocal { param( [string]$myFile, [bool]$HeaderInFirstRow )

if ( -not (Test-Path $myFile)) {
    "not found"
} else {
    

    if ($myFile.Size -gt 0) {
        if ($HeaderInFirstRow) {
            $result = "test"#Check-ContainsMultipleLines -FileName $myFile
            if ($result) {
                "OK"
            } else {
                "Is empty"
            }
        } else {
            "OK"
        }
    } else {
        "Is empty"
    }
}
}

$answer =  Check-FileExistsLocal $myFile, $HeaderInFirstRow

Write-Host $answer

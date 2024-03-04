$testas ="testaresadaasdddddddddsd"
$test = $([System.IO.Path]::GetFileName($testas)).Substring(0,$([System.IO.Path]::GetFileName($testas)).Lenght -5)

Write-Host $test

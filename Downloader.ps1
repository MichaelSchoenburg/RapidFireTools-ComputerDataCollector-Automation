$Url = "https://s3.amazonaws.com/networkdetective/download/NetworkDetectiveComputerDataCollector.exe"
$CurrentLocation = [System.IO.FileInfo]((Get-Location).Path)
$DestPath = "$( $CurrentLocation.FullName )\NetworkDetectiveComputerDataCollector.exe"
$DestPathZip = "$( $CurrentLocation.FullName )\NetworkDetectiveComputerDataCollector.zip"

$Client = New-Object System.Net.WebClient
$Client.DownloadFile($Url, $DestPath)

Move-Item $DestPath $DestPathZip
Expand-Archive -Path $DestPathZip
Remove-Item -Path $DestPathZip

Write-Host "Fertig. Konsole kann geschlossen werden." -ForegroundColor Green
Read-Host

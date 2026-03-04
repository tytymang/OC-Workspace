
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$downloadsPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads")
Get-ChildItem -Path $downloadsPath | Where-Object { $_.Name -like "*식단*" } | Select-Object Name, LastWriteTime | ConvertTo-Json

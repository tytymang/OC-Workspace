
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$downloadsPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads")
$files = Get-ChildItem -Path $downloadsPath | Where-Object { $_.Name -match "2\.23" }
$files | Select-Object Name, LastWriteTime | ConvertTo-Json


[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$downloadsPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads")
$files = Get-ChildItem -Path $downloadsPath | Where-Object { $_.Name -like "*식단*" -or $_.Name -like "*2.23*" }
foreach ($f in $files) {
    Write-Output "File: $($f.Name) | Size: $($f.Length) | LastWrite: $($f.LastWriteTime)"
}

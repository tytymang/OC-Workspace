
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Outlook = New-Object -ComObject Outlook.Application
$Calendar = $Outlook.GetNamespace("MAPI").GetDefaultFolder(9)
$Calendar.Items | Where-Object { $_.Start -gt (Get-Date "2026-04-01") -and $_.Subject -like "*출근*" } | Select-Object Subject, Start | ConvertTo-Json

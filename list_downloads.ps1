
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$downloadsPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads")
# 모든 파일을 날짜순으로 정렬하여 최근 20개만 확인
Get-ChildItem -Path $downloadsPath | Sort-Object LastWriteTime -Descending | Select-Object -First 20 | Select-Object Name, LastWriteTime | ConvertTo-Json

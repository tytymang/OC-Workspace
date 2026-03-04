
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$targetFile = "서울반도체 중석식_2.23 (1).pdf"
$downloadsPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads")
$fullPath = Join-Path $downloadsPath $targetFile
$workspacePath = "C:\Users\307984\.openclaw\workspace\menu.pdf"

if (Test-Path $fullPath) {
    Copy-Item -Path $fullPath -Destination $workspacePath -Force
    Write-Output "SUCCESS: $fullPath"
} else {
    Write-Output "NOT_FOUND: $fullPath"
}

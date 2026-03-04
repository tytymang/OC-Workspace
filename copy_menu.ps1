
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$filePath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads", "서울반도체 중석식_2.23 (1).pdf")
$workspacePath = "C:\Users\307984\.openclaw\workspace\menu.pdf"

if (Test-Path $filePath) {
    Copy-Item $filePath $workspacePath
    Write-Output "SUCCESS"
} else {
    Write-Output "FILE_NOT_FOUND"
}

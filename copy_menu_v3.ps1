
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$downloadsPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads")
$workspacePath = "C:\Users\307984\.openclaw\workspace\menu.pdf"

# 이름에 '2.23'과 '(1)'이 포함된 파일을 찾아 복사합니다.
$file = Get-ChildItem -Path $downloadsPath | Where-Object { $_.Name -like "*2.23*(1)*" -and $_.Extension -eq ".pdf" } | Select-Object -First 1

if ($file) {
    Copy-Item -Path $file.FullName -Destination $workspacePath -Force
    Write-Output "SUCCESS: $($file.Name)"
} else {
    Write-Output "FILE_NOT_FOUND"
}

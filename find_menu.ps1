
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$downloadsPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads")
if (Test-Path $downloadsPath) {
    # '식단'이 포함된 파일 중 2월 23일 전후로 수정된 파일을 찾습니다.
    $files = Get-ChildItem -Path $downloadsPath | Where-Object { $_.Name -like "*식단*" } | Sort-Object LastWriteTime -Descending
    if ($files) {
        $files | Select-Object Name, LastWriteTime, FullName | ConvertTo-Json
    } else {
        Write-Output "FILE_NOT_FOUND"
    }
} else {
    Write-Output "DOWNLOADS_PATH_NOT_FOUND"
}

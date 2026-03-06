# Korean Encoding Skill Applied
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 임시 폴더 경로
$tempPath = "C:\Users\307984\.openclaw\workspace\temp_rates"

# 파일 찾기 (영어 패턴 사용!)
$targetFile = Get-ChildItem -Path $tempPath -Filter "*FY2026*.xlsx" | Select-Object -First 1

if (!$targetFile) {
    Write-Error "Rate file not found (FY2026 pattern)."
    exit
}

$path = $targetFile.FullName
Write-Host "Opening: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    
    # 2025 시트 찾기
    $sheet = $workbook.Sheets.Item("2025")
    Write-Host "Sheet: $($sheet.Name)"
    
    # 데이터 추출 (C열: 기말환율, B열: 월)
    $jan26 = $sheet.Cells.Item(6, 3).Text
    $feb26 = $sheet.Cells.Item(7, 3).Text
    
    # 12월 찾기 (B열 스캔)
    $dec25 = ""
    for ($r = 30; $r -le 40; $r++) {
        $val = $sheet.Cells.Item($r, 2).Text
        if ($val -match "12") {
            $dec25 = $sheet.Cells.Item($r, 3).Text
            break
        }
    }

    Write-Host "DEC_2025: $dec25"
    Write-Host "JAN_2026: $jan26"
    Write-Host "FEB_2026: $feb26"

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

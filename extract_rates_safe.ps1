# Korean Encoding Skill Applied
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 임시 폴더 경로 (한글 없음)
$tempPath = "C:\Users\307984\.openclaw\workspace\temp_rates"
$fileName = "FY2026_Rates.xlsx" # 찾아서 이 이름으로 간주

# 파일 찾기 (temp_rates 폴더 내)
$targetFile = Get-ChildItem -Path $tempPath -Filter "*환율*.xlsx" | Select-Object -First 1

if (!$targetFile) {
    Write-Error "Rate file not found in temp folder."
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
    # 2026년 1월 (Row 6)
    $jan26 = $sheet.Cells.Item(6, 3).Text
    # 2026년 2월 (Row 7)
    $feb26 = $sheet.Cells.Item(7, 3).Text
    # 2025년 12월 (Row 35 근처 - 확인 필요)
    
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

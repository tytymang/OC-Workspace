# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 절대 경로 사용 (경로 문제 원천 차단)
$path = "C:\Users\307984\.openclaw\workspace\temp_exchange.xlsx"
Write-Host "Trying absolute path: $path"

if (!(Test-Path $path)) {
    Write-Error "File still not found at absolute path: $path"
    # 혹시 working 아래에 복사되었는지 확인
    $altPath = "C:\Users\307984\.openclaw\workspace\working\temp_exchange.xlsx"
    if (Test-Path $altPath) {
        $path = $altPath
        Write-Host "Found at alt path: $path"
    } else {
        exit
    }
}

try {
    $workbook = $excel.Workbooks.Open($path)
    $sheet = $workbook.Sheets.Item(1) # 첫 번째 시트

    # 데이터 읽기
    $usedRange = $sheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count

    Write-Host "Rows: $rowCount"

    # 헤더 찾기 (1~5행 스캔)
    $dateCol = 0
    $usdCol = 0

    for ($r = 1; $r -le 5; $r++) {
        for ($c = 1; $c -le $colCount; $c++) {
            $val = $sheet.Cells.Item($r, $c).Text
            if ($val -match "USD|미국|달러") { $usdCol = $c }
            if ($val -match "Date|Month|기간|년월") { $dateCol = $c }
        }
    }
    
    if ($dateCol -eq 0) { $dateCol = 1 }
    if ($usdCol -eq 0) { $usdCol = 2 }

    Write-Host "Target Columns: Date=$dateCol, USD=$usdCol"

    # 데이터 추출
    $targets = @("2025.12", "2026.01", "2026.02", "2025-12", "2026-01", "2026-02", "Dec-25", "Jan-26", "Feb-26")
    
    for ($r = 1; $r -le $rowCount; $r++) {
        $d = $sheet.Cells.Item($r, $dateCol).Text
        $u = $sheet.Cells.Item($r, $usdCol).Text
        
        # 1. 단순 포함
        foreach ($t in $targets) {
            if ($d -like "*$t*") {
                Write-Host "RESULT: $d => $u"
            }
        }
        
        # 2. 정규식 (25.12, 26.01)
        if ($d -match "25\.12" -or $d -match "26\.0?1" -or $d -match "26\.0?2") {
             Write-Host "RESULT(Short): $d => $u"
        }
        
        # 3. 년월 형식 (2025년 12월 등)
        if ($d -match "2025.*12" -or $d -match "2026.*1" -or $d -match "2026.*2") {
             Write-Host "RESULT(Ko): $d => $u"
        }
    }

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

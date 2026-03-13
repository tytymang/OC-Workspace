# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 파일 경로 (절대 경로로 변환)
$path = Convert-Path "working\20260306_재고 꼬리표 작업\통합_FY2026_환율표_202602.xlsx"
$workbook = $excel.Workbooks.Open($path)
$sheet = $workbook.Sheets.Item(1) # 첫 번째 시트라고 가정

# 데이터 읽기 (전체 범위 스캔)
$usedRange = $sheet.UsedRange
$rowCount = $usedRange.Rows.Count
$colCount = $usedRange.Columns.Count

# 결과 저장용 리스트
$results = @()

# 헤더 찾기 (날짜/년월 및 USD 컬럼)
$dateCol = 0
$usdCol = 0

# 1행부터 5행까지 헤더 스캔
for ($r = 1; $r -le 5; $r++) {
    for ($c = 1; $c -le $colCount; $c++) {
        $val = $sheet.Cells.Item($r, $c).Text
        if ($val -match "년월|Date|Month|기간") { $dateCol = $c }
        if ($val -match "USD|미국|달러") { $usdCol = $c }
    }
    if ($dateCol -gt 0 -and $usdCol -gt 0) { break }
}

if ($dateCol -eq 0) { $dateCol = 1 } # 못 찾으면 1열 가정
if ($usdCol -eq 0) { $usdCol = 2 }   # 못 찾으면 2열 가정

Write-Output "Header Found: DateCol=$dateCol, UsdCol=$usdCol"

# 데이터 스캔
for ($r = 1; $r -le $rowCount; $r++) {
    $dateVal = $sheet.Cells.Item($r, $dateCol).Text
    $usdVal = $sheet.Cells.Item($r, $usdCol).Text
    
    # 2025.12, 2026.01, 2026.02 등 다양한 포맷 매칭
    # 2025-12, 202512 등
    
    if ($dateVal -match "2025.*12|2025\.12|Dec.*25" -or $dateVal -match "2026.*0?1|2026\.0?1|Jan.*26" -or $dateVal -match "2026.*0?2|2026\.0?2|Feb.*26") {
         $results += "$dateVal : $usdVal"
    }
}

# 결과 출력
$results

# 정리
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

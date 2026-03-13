# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 절대 경로로 명확히 지정 (현재 위치 기준)
$currentDir = Get-Location
$path = Join-Path $currentDir.Path "working\20260306_재고 꼬리표 작업\통합_FY2026_환율표_202602.xlsx"

Write-Host "Target File: $path"

if (!(Test-Path $path)) {
    Write-Error "File not found at: $path"
    exit
}

try {
    $workbook = $excel.Workbooks.Open($path)
    $sheet = $workbook.Sheets.Item(1)

    # 데이터 읽기
    $usedRange = $sheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count

    Write-Host "Reading $rowCount rows..."

    # 헤더 찾기 (1~5행 스캔)
    $dateCol = 0
    $usdCol = 0

    for ($r = 1; $r -le 5; $r++) {
        for ($c = 1; $c -le $colCount; $c++) {
            $val = $sheet.Cells.Item($r, $c).Text
            if ($val -match "년월|Date|Month|기간") { $dateCol = $c }
            if ($val -match "USD|미국|달러") { $usdCol = $c }
        }
        if ($dateCol -gt 0 -and $usdCol -gt 0) { break }
    }

    # 못 찾으면 기본값 (A열:날짜, B열:USD)
    if ($dateCol -eq 0) { $dateCol = 1 }
    
    # USD 컬럼 못 찾았으면 데이터 패턴으로 추정
    if ($usdCol -eq 0) { 
        for ($c = 2; $c -le $colCount; $c++) { # 보통 2열부터
             $val = $sheet.Cells.Item(2, $c).Text # 2행 데이터 확인
             if ($val -match "^\d+\.\d+$" -or $val -match "^\d{3,4}$") { # 숫자 형태면 USD 후보
                 $usdCol = $c
                 break
             }
        }
    }
    if ($usdCol -eq 0) { $usdCol = 2 }

    Write-Host "Columns: Date=$dateCol, USD=$usdCol"

    # 데이터 추출 (2025.12, 2026.01, 2026.02)
    # 정규식 패턴: 2025.12, 25.12, Dec-25 등
    
    # 2025년 12월 ~ 2026년 2월
    $targets = @("2025.12", "2026.01", "2026.02", "2025-12", "2026-01", "2026-02")
    
    for ($r = 1; $r -le $rowCount; $r++) {
        $d = $sheet.Cells.Item($r, $dateCol).Text
        $u = $sheet.Cells.Item($r, $usdCol).Text
        
        # 1. 단순 포함 여부
        foreach ($t in $targets) {
            if ($d -like "*$t*") {
                Write-Host "FOUND: $d => $u (USD)"
            }
        }
        
        # 2. 정규식 매칭 (한글 포함 등 복잡한 경우)
        if ($d -match "2025.*12" -or $d -match "2026.*0?1" -or $d -match "2026.*0?2") {
             Write-Host "MATCH: $d => $u (USD)"
        }
    }

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

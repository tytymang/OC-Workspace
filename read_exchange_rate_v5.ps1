# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 1. 파일 찾기 (영어 패턴 사용)
# FY2026이 포함된 모든 xlsx 파일 검색
$targetFile = Get-ChildItem -Path "working" -Recurse -Filter "*FY2026*.xlsx" | Select-Object -First 1

if (!$targetFile) {
    Write-Error "FY2026 파일을 찾을 수 없습니다."
    exit
}

$path = $targetFile.FullName
Write-Host "File Found: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    $sheet = $workbook.Sheets.Item(1)

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
    
    # 못 찾으면 기본값
    if ($dateCol -eq 0) { $dateCol = 1 }
    if ($usdCol -eq 0) { $usdCol = 2 }

    Write-Host "Target Columns: Date=$dateCol, USD=$usdCol"

    # 데이터 추출
    # 2025.12, 2026.01, 2026.02
    
    $targets = @("2025.12", "2026.01", "2026.02", "2025-12", "2026-01", "2026-02", "Dec-25", "Jan-26", "Feb-26")
    
    for ($r = 1; $r -le $rowCount; $r++) {
        $d = $sheet.Cells.Item($r, $dateCol).Text
        $u = $sheet.Cells.Item($r, $usdCol).Text
        
        foreach ($t in $targets) {
            if ($d -like "*$t*") {
                Write-Host "RESULT: $d => $u"
            }
        }
        
        # 25.12, 26.01 등 짧은 형식도 체크
        if ($d -match "25\.12" -or $d -match "26\.01" -or $d -match "26\.02") {
            Write-Host "RESULT(Short): $d => $u"
        }
    }

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

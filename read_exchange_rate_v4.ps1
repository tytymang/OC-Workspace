# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 1. 파일 찾기 (와일드카드로 검색하여 한글 경로 문제 우회)
# working 폴더 아래에서 '환율'이 포함된 xlsx 파일 검색
$targetFile = Get-ChildItem -Path "working" -Recurse -Filter "*환율*.xlsx" | Select-Object -First 1

if (!$targetFile) {
    Write-Error "환율표 파일을 찾을 수 없습니다."
    exit
}

$path = $targetFile.FullName
Write-Host "Found File: $path"

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

    # 못 찾으면 기본값
    if ($dateCol -eq 0) { $dateCol = 1 }
    if ($usdCol -eq 0) { 
        # 데이터로 추정
        for ($c = 2; $c -le $colCount; $c++) { 
             $val = $sheet.Cells.Item(6, $c).Text # 6행 데이터 확인 (헤더 이후)
             if ($val -match "^\d+\.?\d*") { 
                 $usdCol = $c
                 break
             }
        }
    }
    if ($usdCol -eq 0) { $usdCol = 2 }

    Write-Host "Columns: Date=$dateCol, USD=$usdCol"

    # 데이터 추출 (2025.12, 2026.01, 2026.02)
    # 2025년 12월 ~ 2026년 2월
    
    for ($r = 1; $r -le $rowCount; $r++) {
        $d = $sheet.Cells.Item($r, $dateCol).Text
        $u = $sheet.Cells.Item($r, $usdCol).Text
        
        # 날짜 포맷 확인 (다양한 형식 대응)
        # 2025-12, 2025.12, 202512
        # 2026-01, 2026.01, 202601
        # 2026-02, 2026.02, 202602
        
        if ($d -match "2025.*12" -or $d -match "2026.*0?1" -or $d -match "2026.*0?2") {
            Write-Host "RESULT: $d => $u"
        }
    }

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

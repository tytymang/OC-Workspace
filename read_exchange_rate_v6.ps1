# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 1. 짧은 경로 사용
$shortPath = "working\202603~1"
$fullPath = Resolve-Path $shortPath
Write-Host "Target Dir: $fullPath"

# 2. 파일 찾기 (영어 패턴 사용)
# FY2026이 포함된 모든 xlsx 파일 검색
$targetFile = Get-ChildItem -Path $shortPath -Filter "*FY2026*.xlsx" | Select-Object -First 1

if (!$targetFile) {
    Write-Error "FY2026 파일을 찾을 수 없습니다."
    exit
}

# 3. 경로 재조합 (한글 깨짐 방지 위해 짧은 경로 사용)
# 파일명은 한글이 포함될 수 있으므로, 파일명도 8.3 형식이 있으면 좋지만,
# Get-ChildItem은 기본적으로 FullName을 반환하므로, 
# 여기서는 파일명만 가져와서 수동으로 경로를 조합한다.

$fileName = $targetFile.Name
# 만약 파일명에도 한글이 있어 문제가 된다면, 와일드카드로 첫번째 파일을 연다.
# 하지만 일단 시도.

# 절대 경로로 변환 (현재 위치 기준)
$currentDir = Get-Location
# 파일명을 포함한 전체 경로 (짧은 경로 사용)
$filePath = Join-Path $currentDir.Path "$shortPath\$fileName"

Write-Host "Trying to open: $filePath"

try {
    $workbook = $excel.Workbooks.Open($filePath)
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
    
    if ($dateCol -eq 0) { $dateCol = 1 }
    if ($usdCol -eq 0) { 
        # 데이터로 추정
        for ($c = 2; $c -le $colCount; $c++) { 
             $val = $sheet.Cells.Item(6, $c).Text 
             if ($val -match "^\d+\.?\d*") { 
                 $usdCol = $c
                 break
             }
        }
    }
    if ($usdCol -eq 0) { $usdCol = 2 }

    Write-Host "Cols: Date=$dateCol, USD=$usdCol"

    # 데이터 추출
    $targets = @("2025.12", "2026.01", "2026.02", "2025-12", "2026-01", "2026-02")
    
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
            # 중복 방지 로직은 생략 (단순 출력)
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

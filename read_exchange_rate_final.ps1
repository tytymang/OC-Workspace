# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 1. temp_work 폴더에서 파일 찾기
# 한글 폴더가 있어도 Recurse가 동작하는지 확인
$targetFile = Get-ChildItem -Path "temp_work" -Recurse -Filter "*FY2026*.xlsx" | Select-Object -First 1

if (!$targetFile) {
    Write-Error "FY2026 파일을 찾을 수 없습니다."
    # 만약 Recurse 실패 시, 직접 경로 지정 시도
    # 하지만 폴더명이 한글이라...
    # Robocopy로 Flatten Copy를 했어야 했는데, 구조 그대로 복사됨.
    
    # 최후의 수단: temp_work 안의 모든 폴더를 뒤져서 파일 찾기 (Wildcard)
    $subDirs = Get-ChildItem -Path "temp_work" -Directory
    foreach ($d in $subDirs) {
        $f = Get-ChildItem -Path $d.FullName -Filter "*FY2026*.xlsx" | Select-Object -First 1
        if ($f) {
            $targetFile = $f
            break
        }
    }
}

if (!$targetFile) {
    Write-Error "Still not found."
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
    
    if ($dateCol -eq 0) { $dateCol = 1 }
    if ($usdCol -eq 0) { $usdCol = 2 } # 기본값

    Write-Host "Cols: Date=$dateCol, USD=$usdCol"

    # 데이터 추출
    $targets = @("2025.12", "2026.01", "2026.02", "2025-12", "2026-01", "2026-02")
    
    for ($r = 1; $r -le $rowCount; $r++) {
        $d = $sheet.Cells.Item($r, $dateCol).Text
        $u = $sheet.Cells.Item($r, $usdCol).Text
        
        foreach ($t in $targets) {
            if ($d -like "*$t*") {
                Write-Host "RESULT: $d => $u"
            }
        }
        
        if ($d -match "25\.12" -or $d -match "26\.0?1" -or $d -match "26\.0?2") {
             Write-Host "RESULT(Short): $d => $u"
        }
        
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
    
    # 임시 폴더 삭제
    Remove-Item "temp_work" -Recurse -Force
}

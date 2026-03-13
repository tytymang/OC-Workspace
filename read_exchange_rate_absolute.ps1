# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 절대 경로 사용 (경로 문제 원천 차단)
$baseDir = "C:\Users\307984\.openclaw\workspace\temp_work"

if (!(Test-Path $baseDir)) {
    Write-Error "Temp work directory not found at: $baseDir"
    exit
}

# 1. 파일 찾기 (절대 경로 기준)
$targetFile = Get-ChildItem -Path $baseDir -Recurse -Filter "*FY2026*.xlsx" | Select-Object -First 1

if (!$targetFile) {
    # 혹시 한글 폴더 때문에 못 찾으면, 직접 순회
    $subDirs = Get-ChildItem -Path $baseDir -Directory
    foreach ($d in $subDirs) {
        $f = Get-ChildItem -Path $d.FullName -Filter "*FY2026*.xlsx" | Select-Object -First 1
        if ($f) {
            $targetFile = $f
            break
        }
    }
}

if (!$targetFile) {
    Write-Error "File not found inside temp_work."
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
    if ($usdCol -eq 0) { $usdCol = 2 }

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
    Remove-Item $baseDir -Recurse -Force
}

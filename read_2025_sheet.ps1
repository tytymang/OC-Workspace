# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 절대 경로 사용
$baseDir = "C:\Users\307984\.openclaw\workspace\temp_work"

# 파일 찾기
$targetFile = Get-ChildItem -Path $baseDir -Recurse -Filter "*FY2026*.xlsx" | Select-Object -First 1

if (!$targetFile) {
    # 서브디렉토리 순회
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
    Write-Error "File not found."
    exit
}

$path = $targetFile.FullName
Write-Host "Target: 2025 Sheet in $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    
    # 2025 시트 찾기
    $sheet = $null
    foreach ($s in $workbook.Sheets) {
        if ($s.Name -eq "2025") {
            $sheet = $s
            break
        }
    }
    
    if (!$sheet) {
        Write-Error "Sheet '2025' not found. Checking all sheets..."
        # 만약 2025 시트가 없으면, 모든 시트를 뒤져서 2025-12가 있는 곳을 찾음
        foreach ($s in $workbook.Sheets) {
             # 각 시트의 1열이나 2열을 빠르게 스캔
             $found = $false
             for ($r = 1; $r -le 20; $r++) {
                 if ($s.Cells.Item($r, 1).Text -match "2025.*12" -or $s.Cells.Item($r, 1).Text -match "2026.*0?1") {
                     $sheet = $s
                     $found = $true
                     break
                 }
             }
             if ($found) { break }
        }
    }
    
    if (!$sheet) {
        Write-Error "Target data not found in any sheet."
        exit
    }
    
    Write-Host "Reading Sheet: $($sheet.Name)"

    $usedRange = $sheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count

    Write-Host "Rows: $rowCount"

    # 헤더 찾기 (USD, Date)
    $dateCol = 0
    $usdCol = 0

    for ($r = 1; $r -le 10; $r++) { # 10행까지 스캔
        for ($c = 1; $c -le $colCount; $c++) {
            $val = $sheet.Cells.Item($r, $c).Text
            if ($val -match "USD") { $usdCol = $c }
            if ($val -match "Date|Month|Year") { $dateCol = $c }
        }
    }
    
    # 헤더 못 찾으면 데이터로 추정
    if ($dateCol -eq 0) { 
        # 날짜 형식(2025.12 등)이 있는 열 찾기
        for ($c = 1; $c -le $colCount; $c++) {
            for ($r = 1; $r -le 20; $r++) {
                if ($sheet.Cells.Item($r, $c).Text -match "2025.*12") {
                    $dateCol = $c
                    break
                }
            }
            if ($dateCol -gt 0) { break }
        }
    }
    
    if ($usdCol -eq 0) { 
        # USD 컬럼 추정 (Date 컬럼 옆이나, 1300~1500 사이 숫자)
        for ($c = 1; $c -le $colCount; $c++) {
            for ($r = 5; $r -le 20; $r++) {
                $val = $sheet.Cells.Item($r, $c).Text
                # 1000 ~ 2000 사이 숫자
                if ($val -match "^1[0-9]{3}") {
                    $usdCol = $c
                    break
                }
            }
             if ($usdCol -gt 0) { break }
        }
    }

    Write-Host "Cols: Date=$dateCol, USD=$usdCol"
    
    if ($dateCol -eq 0) { $dateCol = 1 }
    if ($usdCol -eq 0) { $usdCol = 2 }

    # 데이터 추출
    $targets = @("2025.12", "2026.01", "2026.02", "2025-12", "2026-01", "2026-02")
    
    # 2026년 2월이 없을 수도 있으니, 2025년 12월부터 최대한 찾기
    
    for ($r = 1; $r -le $rowCount; $r++) {
        $d = $sheet.Cells.Item($r, $dateCol).Text
        $u = $sheet.Cells.Item($r, $usdCol).Text
        
        # 영어/숫자 패턴만 사용
        if ($d -match "2025.*12" -or $d -match "2026.*0?1" -or $d -match "2026.*0?2") {
             Write-Host "RESULT: $d => $u"
        }
        
        # 12월, 1월, 2월 숫자만 있는 경우도 대비 (예: 12월)
        if ($d -eq "12" -or $d -eq "1" -or $d -eq "2" -or $d -eq "01" -or $d -eq "02") {
             # 연도 확인이 필요하지만 일단 출력
             Write-Host "CANDIDATE: $d => $u"
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

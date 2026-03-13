# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 절대 경로 사용
$baseDir = "C:\Users\307984\.openclaw\workspace\temp_work"

# 파일 찾기
$targetFile = Get-ChildItem -Path $baseDir -Recurse -Filter "*FY2026*.xlsx" | Select-Object -First 1

if (!$targetFile) {
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
Write-Host "Open: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    $sheet = $workbook.Sheets.Item(1)

    # 데이터 읽기
    $usedRange = $sheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count

    Write-Host "Rows: $rowCount"

    # 헤더 찾기 (1~5행 스캔) - 영어만 사용!
    $dateCol = 0
    $usdCol = 0

    for ($r = 1; $r -le 5; $r++) {
        for ($c = 1; $c -le $colCount; $c++) {
            $val = $sheet.Cells.Item($r, $c).Text
            # USD, Date 등 영어만 매칭
            if ($val -match "USD") { $usdCol = $c }
            if ($val -match "Date|Month|Year") { $dateCol = $c }
        }
    }
    
    # 못 찾으면 강제 할당
    if ($dateCol -eq 0) { $dateCol = 1 } # A열
    
    if ($usdCol -eq 0) { 
        # 2행 데이터가 숫자면 USD로 추정
        for ($c = 2; $c -le $colCount; $c++) { 
             $val = $sheet.Cells.Item(6, $c).Text 
             if ($val -match "^\d+\.?\d*") { 
                 $usdCol = $c
                 break
             }
        }
    }
    
    if ($usdCol -eq 0) { $usdCol = 2 } # B열

    Write-Host "Cols: Date=$dateCol, USD=$usdCol"

    # 데이터 추출
    $targets = @("2025.12", "2026.01", "2026.02", "2025-12", "2026-01", "2026-02")

    for ($r = 1; $r -le $rowCount; $r++) {
        # 인덱스 안전장치 (1 이상이어야 함)
        if ($dateCol -lt 1) { $dateCol = 1 }
        if ($usdCol -lt 1) { $usdCol = 2 }

        $d = $sheet.Cells.Item($r, $dateCol).Text
        $u = $sheet.Cells.Item($r, $usdCol).Text
        
        # 영어/숫자 패턴만 사용
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
    
    # 임시 폴더 삭제
    Remove-Item $baseDir -Recurse -Force
}

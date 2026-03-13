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
Write-Host "Searching 2025 Header: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    $sheet = $workbook.Sheets.Item("2025")
    
    $usedRange = $sheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    
    Write-Host "Total Rows: $rowCount"
    
    # 2025년 헤더 찾기 (전체 스캔)
    $year2025Row = 0
    
    for ($r = 1; $r -le $rowCount; $r++) {
        # 2열(B열)이나 3열(C열)에서 연도 확인
        $val = $sheet.Cells.Item($r, 2).Text
        $val2 = $sheet.Cells.Item($r, 3).Text
        
        if ($val -match "2025" -or $val2 -match "2025") {
            Write-Host "Found 2025 Header at Row ${r}: $val $val2"
            $year2025Row = $r
            # 여기서 멈추지 않고 계속 찾음 (여러 개일 수 있음)
        }
    }
    
    # 만약 찾았다면 그 아래 12월 데이터 찾기
    if ($year2025Row -gt 0) {
        Write-Host "Scanning below Row ${year2025Row}..."
        for ($r = $year2025Row; $r -le $year2025Row + 20; $r++) {
            $monthVal = $sheet.Cells.Item($r, 2).Text # 월 (B열)
            $usdVal = $sheet.Cells.Item($r, 26).Text # USD (Z열)
            
            if ($monthVal -match "12") {
                Write-Host "CANDIDATE: Row ${r} - Month=[$monthVal] USD=[$usdVal]"
            }
        }
    } else {
        # 2025년 헤더를 못 찾았으면, 2024 시트도 확인해봐야 함
        Write-Host "2025 Header not found in '2025' sheet. Checking '2024' sheet..."
        
        $sheet24 = $workbook.Sheets.Item("2024")
        if ($sheet24) {
             # 2024 시트에서 2025년 데이터 찾기 (아마도 상단에?)
             for ($r = 1; $r -le 20; $r++) {
                $val = $sheet24.Cells.Item($r, 2).Text
                 if ($val -match "2025") {
                     Write-Host "Found 2025 in '2024' sheet at Row ${r}"
                     # 그 아래 12월 찾기
                     for ($subR = $r; $subR -le $r + 15; $subR++) {
                         $m = $sheet24.Cells.Item($subR, 2).Text
                         $u = $sheet24.Cells.Item($subR, 26).Text
                         if ($m -match "12") {
                             Write-Host "CANDIDATE (2024 sheet): Month=[$m] USD=[$u]"
                         }
                     }
                 }
             }
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

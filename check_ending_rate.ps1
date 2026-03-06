# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 1. 대상 파일 정의 (temp_check 폴더)
$baseDir = "C:\Users\307984\.openclaw\workspace\temp_check"

# 파일 찾기
$targetFile = Get-ChildItem -Path $baseDir -Recurse -Filter "*FY2026*.xlsx" | Select-Object -First 1

if (!$targetFile) {
    Write-Error "Target file not found."
    exit
}

$path = $targetFile.FullName
Write-Host "Checking for Ending Rate in 2025 Sheet: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    $sheet = $workbook.Sheets.Item("2025")
    
    # 1행 ~ 10행, 1열 ~ 30열 덤프 (헤더 확인)
    for ($r = 1; $r -le 10; $r++) {
        $rowText = ""
        for ($c = 1; $c -le 30; $c++) { 
            $val = $sheet.Cells.Item($r, $c).Text
            # 기말, 마감, End, Final 등의 키워드 확인
            if ($val) {
                $rowText += "[${c}:$val] "
            }
        }
        Write-Host "Row ${r}: $rowText"
    }

    # 데이터 영역 (2025.12 ~ 2026.02) 다시 스캔
    # 기말환율 컬럼을 찾았다고 가정하고 데이터 확인
    # 아까 26열이 USD 평균환율이었으니, 그 옆(27, 28...)이나 앞(25...)에 있을 수 있음.
    
    Write-Host "--- Data Scan (Dec 2025 ~ Feb 2026) ---"
    
    # 2025년 12월 위치 찾기 (아까 35행 근처였음)
    # 2026년 1월, 2월 위치 찾기 (아까 6, 7행 근처였음)
    
    $rowsToCheck = @(6, 7, 35) # 예상 행
    
    foreach ($r in $rowsToCheck) {
        $dateVal = $sheet.Cells.Item($r, 2).Text # 날짜 (B열)
        $val26 = $sheet.Cells.Item($r, 26).Text # USD 평균?
        $val27 = $sheet.Cells.Item($r, 27).Text # USD 기말? (추정)
        $val14 = $sheet.Cells.Item($r, 14).Text # 다른 USD?
        
        Write-Host "Row ${r}: Date=[$dateVal] Col14=[$val14] Col26=[$val26] Col27=[$val27]"
    }

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# Korean Encoding Skill Applied
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 절대 경로 설정
$basePath = "C:\Users\307984\.openclaw\workspace\working\20260306_재고 꼬리표 작업"
$fileName = "재무_FY2026_환율표_202602.xlsx"
$filePath = Join-Path $basePath $fileName

if (!(Test-Path $filePath)) {
    Write-Error "File not found: $filePath"
    exit
}

Write-Host "Opening: $fileName"

try {
    $workbook = $excel.Workbooks.Open($filePath)
    
    # 시트 찾기: 2025, FY2026, 또는 가장 최신 시트
    $targetSheet = $null
    foreach ($s in $workbook.Sheets) {
        if ($s.Name -match "2025" -or $s.Name -match "FY26" -or $s.Name -match "FY2026") {
            $targetSheet = $s
            break
        }
    }
    
    # 없으면 첫 번째 시트 사용 (보통 최신 데이터가 첫 시트에 있음)
    if (!$targetSheet) { $targetSheet = $workbook.Sheets.Item(1) }
    
    Write-Host "Reading Sheet: $($targetSheet.Name)"
    
    $usedRange = $targetSheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count

    # 헤더 분석 (기말 환율 컬럼 찾기)
    $dateCol = 0
    $usdEndingCol = 0
    
    # 상단 10행 스캔
    for ($r = 1; $r -le 10; $r++) {
        for ($c = 1; $c -le $colCount; $c++) {
            $val = $targetSheet.Cells.Item($r, $c).Text
            
            # 날짜 컬럼 찾기
            if ($val -match "Date" -or $val -match "Month" -or $val -match "기간" -or $val -match "년월") {
                $dateCol = $c
            }
            
            # 기말 환율 컬럼 찾기 (USD + 기말/Ending/Close)
            # 보통 USD라고만 적혀있고 상위 헤더에 '기말'이 있는 경우가 많음
            # 여기서는 '기말' 키워드가 있는 열이나, 그 아래 USD를 찾음
            if ($val -match "기말" -or $val -match "Ending") {
                # 이 열이거나, 이 그룹의 USD 열을 찾아야 함.
                # 단순하게 이 열 아래에 숫자가 있으면 이 열을 USD 기말이라고 가정 (임시)
                # 더 정확히는 행/열 구조를 봐야 함. 
                # 일단 USD 키워드도 같이 확인
            }
            
            if ($val -eq "USD" -or $val -match "미국") {
                 # 상위 행(r-1, r-2)에 '기말'이 있는지 확인하면 좋음
                 # 일단 후보군으로 저장. 보통 왼쪽이 기말, 오른쪽이 평균인 경우가 많음.
                 if ($usdEndingCol -eq 0) { $usdEndingCol = $c } # 첫번째 나오는 USD를 기말로 가정 (보통 관례)
            }
        }
    }
    
    # 컬럼 못 찾았으면 기본값 (A=날짜, C=USD기말 가정 - 이전 경험 기반)
    if ($dateCol -eq 0) { $dateCol = 2 } # B열 (보통 월)
    if ($usdEndingCol -eq 0) { $usdEndingCol = 3 } # C열 (보통 첫번째 환율이 기말)

    Write-Host "Assuming Date Col: $dateCol, USD Ending Rate Col: $usdEndingCol"
    
    # 데이터 추출 (2025.12, 2026.01, 2026.02)
    $targets = @("2025", "2026")
    
    Write-Host "`n--- Extraction Results ---"
    for ($r = 1; $r -le $rowCount; $r++) {
        $dateVal = $targetSheet.Cells.Item($r, $dateCol).Text
        $rateVal = $targetSheet.Cells.Item($r, $usdEndingCol).Text
        
        # 날짜 매칭 (12월, 1월, 2월)
        if ($dateVal -match "12" -or $dateVal -match "01" -or $dateVal -match "02") {
            # 연도 체크 (2025, 2026이 포함되어 있거나, 시트가 2025년 시트라면 월만 봐도 됨)
            # 문맥상 1400원대 환율이 나오면 출력
            if ($rateVal -match "[0-9,]+\.[0-9]+") {
                 Write-Host "Row $r : Date=[$dateVal] Rate=[$rateVal]"
            }
        }
    }

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

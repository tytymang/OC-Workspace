# 엑셀 COM 객체 생성 (Visible=False, Alert=False)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 1. 대상 파일 정의 (02월 보고서)
# 절대 경로 사용 (경로 문제 원천 차단)
$baseDir = "C:\Users\307984\.openclaw\workspace\working\20260306_재고 꼬리표 작업"
$reportFile = "20260305_02월 재고 꼬리표 Report.xlsx"
$path = Join-Path $baseDir $reportFile

# 파일 존재 확인
if (!(Test-Path $path)) {
    Write-Error "Report file not found: $path"
    exit
}

Write-Host "Opening Report: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    $sheet = $workbook.Sheets.Item(1) # 첫 번째 시트 (보통 Summary나 Main)

    # 2. 환율 업데이트 로직
    # 환율 셀 위치를 찾아야 함 (보통 상단이나 특정 위치)
    # "환율" 또는 "USD" 키워드 검색 후 값 변경
    
    $usedRange = $sheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count
    
    # 2월 환율: 1,449.32
    $targetRate = 1449.32
    $updated = $false
    
    # 헤더 영역(1~20행) 스캔하여 환율 셀 찾기
    for ($r = 1; $r -le 20; $r++) {
        for ($c = 1; $c -le $colCount; $c++) {
            $val = $sheet.Cells.Item($r, $c).Text
            
            # "환율" 텍스트가 있는 셀 옆이나 아래에 값이 있을 확률 높음
            if ($val -match "환율" -or $val -match "Exchange Rate") {
                # 옆 셀 확인 (Right)
                $nextVal = $sheet.Cells.Item($r, $c+1).Value2
                if ($nextVal -is [double] -or $nextVal -is [int]) {
                    Write-Host "Found Rate Cell at [${r},${c}+1]: Old=$nextVal -> New=$targetRate"
                    $sheet.Cells.Item($r, $c+1).Value2 = $targetRate
                    $updated = $true
                } 
                # 아래 셀 확인 (Below)
                $belowVal = $sheet.Cells.Item($r+1, $c).Value2
                if (!$updated -and ($belowVal -is [double] -or $belowVal -is [int])) {
                     Write-Host "Found Rate Cell at [${r}+1,${c}]: Old=$belowVal -> New=$targetRate"
                     $sheet.Cells.Item($r+1, $c).Value2 = $targetRate
                     $updated = $true
                }
            }
            if ($updated) { break }
        }
        if ($updated) { break }
    }
    
    if (!$updated) {
        Write-Warning "Failed to locate 'Exchange Rate' cell automatically."
        # 강제로 찾기 위해 셀 주소를 물어보거나, 
        # 특정 셀(예: B2, C2 등)을 추정해서 기록할 수도 있으나, 일단 보고.
    } else {
        # 저장
        $workbook.Save()
        Write-Host "Successfully updated exchange rate to $targetRate"
    }

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

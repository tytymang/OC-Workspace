# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 1. 대상 파일 정의 (temp_report 폴더)
$baseDir = "C:\Users\307984\.openclaw\workspace\temp_report"

# 파일 찾기 (한글 경로 문제 해결 위해 파일명 패턴 검색)
$targetFile = Get-ChildItem -Path $baseDir -Recurse -Filter "*Report*.xlsx" | Where-Object { $_.Name -like "*02*" } | Select-Object -First 1

if (!$targetFile) {
    Write-Error "Report file not found in temp_report."
    exit
}

$path = $targetFile.FullName
Write-Host "Updating Report: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    $sheet = $workbook.Sheets.Item(1)

    # 2. 환율 업데이트 로직
    $targetRate = 1449.32
    $updated = $false
    
    $usedRange = $sheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count
    
    # 헤더 스캔
    for ($r = 1; $r -le 20; $r++) {
        for ($c = 1; $c -le $colCount; $c++) {
            $val = $sheet.Cells.Item($r, $c).Text
            
            # 환율 키워드 찾기 (USD, 환율)
            if ($val -match "USD" -or $val -match "환율" -or $val -match "Rate") {
                # 바로 오른쪽 셀 확인
                $nextVal = $sheet.Cells.Item($r, $c+1).Value2
                if ($nextVal -is [double] -or $nextVal -is [int]) {
                    Write-Host "Updating cell [${r},${c}+1]: $nextVal -> $targetRate"
                    $sheet.Cells.Item($r, $c+1).Value2 = $targetRate
                    $updated = $true
                }
                # 바로 아래 셀 확인
                $belowVal = $sheet.Cells.Item($r+1, $c).Value2
                if (!$updated -and ($belowVal -is [double] -or $belowVal -is [int])) {
                     Write-Host "Updating cell [${r}+1,${c}]: $belowVal -> $targetRate"
                     $sheet.Cells.Item($r+1, $c).Value2 = $targetRate
                     $updated = $true
                }
            }
            if ($updated) { break }
        }
        if ($updated) { break }
    }
    
    if ($updated) {
        $workbook.Save()
        Write-Host "SUCCESS: Report updated."
    } else {
        Write-Warning "Could not find exchange rate cell automatically."
    }

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

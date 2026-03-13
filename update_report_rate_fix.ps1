# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 1. 대상 파일 정의 (temp_report 폴더)
$baseDir = "C:\Users\307984\.openclaw\workspace\temp_report"

# 파일 찾기 (02월 보고서만 선택)
# "20260305_02월" 패턴 사용 (파일명에 한글이 있어도 와일드카드로 매칭)
$targetFile = Get-ChildItem -Path $baseDir -Recurse -Filter "*20260305_02*.xlsx" | Select-Object -First 1

if (!$targetFile) {
    Write-Error "Target 02 Report file not found."
    exit
}

$path = $targetFile.FullName
Write-Host "Updating: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    $sheet = $workbook.Sheets.Item(1)

    # 2. 환율 업데이트 로직
    $targetRate = 1449.32
    $updated = $false
    
    $usedRange = $sheet.UsedRange
    $colCount = $usedRange.Columns.Count
    
    # 헤더 스캔 (영어 키워드만 사용!)
    for ($r = 1; $r -le 20; $r++) {
        for ($c = 1; $c -le $colCount; $c++) {
            $val = $sheet.Cells.Item($r, $c).Text
            
            # USD, Rate 등 영어만 매칭
            if ($val -match "USD" -or $val -match "Rate") {
                # 오른쪽 셀 확인
                $nextVal = $sheet.Cells.Item($r, $c+1).Value2
                if ($nextVal -is [double] -or $nextVal -is [int]) {
                    # 기존 값이 1000 이상인 경우만 환율로 간주 (안전장치)
                    if ($nextVal -gt 1000) {
                        Write-Host "Updating Cell [${r},${c}+1]: $nextVal -> $targetRate"
                        $sheet.Cells.Item($r, $c+1).Value2 = $targetRate
                        $updated = $true
                    }
                }
                
                # 아래 셀 확인
                $belowVal = $sheet.Cells.Item($r+1, $c).Value2
                if (!$updated -and ($belowVal -is [double] -or $belowVal -is [int])) {
                     if ($belowVal -gt 1000) {
                         Write-Host "Updating Cell [${r}+1,${c}]: $belowVal -> $targetRate"
                         $sheet.Cells.Item($r+1, $c).Value2 = $targetRate
                         $updated = $true
                     }
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
        Write-Warning "Could not find exchange rate cell automatically. (Checked USD/Rate keywords)"
        # 못 찾았을 경우, 특정 셀(B4, C4 등)을 강제로 확인해볼 수도 있음
    }

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

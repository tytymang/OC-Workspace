# Korean Encoding Skill Applied
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 경로 설정 (temp_test)
$basePath = "C:\Users\307984\.openclaw\workspace\temp_test"
$targetFile = Get-ChildItem -Path $basePath -Recurse -Filter "*01월말*Report*.xlsx" | Select-Object -First 1

if (!$targetFile) { Write-Error "File not found"; exit }
$path = $targetFile.FullName
Write-Host "Editing: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    $found = $false
    $exchRate = 1427.00
    
    # 모든 시트 검색
    foreach ($sheet in $workbook.Sheets) {
        $usedRange = $sheet.UsedRange
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count
        
        # '선봉규' 찾기 (Find 메서드 사용)
        $searchRange = $sheet.Range("A1", $sheet.Cells.Item($rowCount, $colCount))
        $foundCell = $searchRange.Find("선봉규")
        
        if ($foundCell) {
            $row = $foundCell.Row
            Write-Host "Found TARGET in Sheet '$($sheet.Name)' at Row $row"
            
            # 컬럼 매핑 (헤더 검색 - 영어 키워드 사용)
            $krwCol = 0
            $usdCol = 0
            
            # 헤더 행 추정 (데이터 위쪽 1~5행 스캔)
            for ($h = 1; $h -le $row; $h++) {
                for ($c = 1; $c -le $colCount; $c++) {
                    $val = $sheet.Cells.Item($h, $c).Text
                    if ($val -match "Accounting" -or $val -match "Inventory" -or $val -match "KRW") { $krwCol = $c }
                    if ($val -match "USD") { $usdCol = $c }
                }
                if ($krwCol -gt 0) { break }
            }
            
            # 못 찾았으면 수동 매핑 (보통 P열이 KRW, N열이 USD였음)
            if ($krwCol -eq 0) { $krwCol = 16 } # P열 가정
            if ($usdCol -eq 0) { $usdCol = 14 } # N열 가정
            
            # 값 읽기 (억원 단위)
            $krwVal = $sheet.Cells.Item($row, $krwCol).Value2
            
            if ($krwVal -is [double] -or $krwVal -is [int]) {
                # 계산: (억원 * 1억) / 환율
                $realKRW = $krwVal * 100000000
                $usdVal = $realKRW / $exchRate
                
                Write-Host "Original KRW (100M Unit): $krwVal"
                Write-Host "Real KRW: $realKRW"
                Write-Host "Calculated USD: $usdVal (Rate: $exchRate)"
                
                if ($usdCol -gt 0) {
                    $sheet.Cells.Item($row, $usdCol).Value2 = $usdVal
                    Write-Host "Updated Cell [Row ${row}, Col ${usdCol}] with USD value."
                } else {
                    $sheet.Cells.Item($row, $krwCol + 1).Value2 = $usdVal
                    Write-Host "Wrote to Right Cell [Row ${row}, Col $($krwCol+1)]"
                }
                $found = $true
            } else {
                Write-Warning "KRW value is not a number: $krwVal"
            }
            break # 한 명만 수정
        }
    }
    
    if ($found) {
        $workbook.Save()
        Write-Host "SUCCESS: Saved changes."
    } else {
        Write-Warning "Target not found."
    }

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

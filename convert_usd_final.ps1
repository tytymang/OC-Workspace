# Korean Encoding Skill Applied
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 경로 설정 (temp_test)
$basePath = "C:\Users\307984\.openclaw\workspace\temp_test"
$targetFile = Get-ChildItem -Path $basePath -Recurse -Filter "*01월말*Report*.xlsx" | Select-Object -First 1

if (!$targetFile) { exit }
$path = $targetFile.FullName

try {
    $workbook = $excel.Workbooks.Open($path)
    $found = $false
    $exchRate = 1427.00
    
    foreach ($sheet in $workbook.Sheets) {
        $usedRange = $sheet.UsedRange
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count
        
        $searchRange = $sheet.Range("A1", $sheet.Cells.Item($rowCount, $colCount))
        $foundCell = $searchRange.Find("선봉규")
        
        if ($foundCell) {
            $row = $foundCell.Row
            
            $krwCol = 0
            $usdCol = 0
            
            for ($h = 1; $h -le $row; $h++) {
                for ($c = 1; $c -le $colCount; $c++) {
                    $val = $sheet.Cells.Item($h, $c).Text
                    if ($val -match "Accounting" -or $val -match "Inventory" -or $val -match "KRW") { $krwCol = $c }
                    if ($val -match "USD") { $usdCol = $c }
                }
                if ($krwCol -gt 0) { break }
            }
            
            if ($krwCol -eq 0) { $krwCol = 16 }
            if ($usdCol -eq 0) { $usdCol = 14 }
            
            $krwVal = $sheet.Cells.Item($row, $krwCol).Value2
            
            if ($krwVal -is [double] -or $krwVal -is [int]) {
                $realKRW = $krwVal * 100000000
                $usdVal = $realKRW / $exchRate
                
                # 출력 대신 변수에 저장해뒀다가 마지막에 확인
                $sheet.Cells.Item($row, $usdCol).Value2 = $usdVal
                $found = $true
            }
            break
        }
    }
    
    if ($found) {
        $workbook.Save()
        # 성공 메시지 단순화
        Write-Output "DONE"
    }

} catch {
    # 에러 무시
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

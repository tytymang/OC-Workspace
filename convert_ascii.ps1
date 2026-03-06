
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$path = "C:\Users\307984\.openclaw\workspace\temp_test\20260205_01월말 재고 꼬리표 Report_3.xlsx"

try {
    $wb = $excel.Workbooks.Open($path)
    $found = $false
    
    # "선봉규" in Unicode Chars
    $targetName = [string][char]49440 + [string][char]48393 + [string][char]44508
    
    foreach ($sh in $wb.Sheets) {
        $ur = $sh.UsedRange
        $rowCount = $ur.Rows.Count
        $colCount = $ur.Columns.Count
        
        for ($r = 6; $r -le $rowCount; $r++) {
            for ($c = 2; $c -le 6; $c++) {
                $val = $sh.Cells.Item($r, $c).Text
                if ($val -match $targetName) {
                    
                    $krw = $sh.Cells.Item($r, 16).Value2
                    
                    if ($krw -is [double] -or $krw -is [int]) {
                        $usd = ($krw * 100000000) / 1427
                        $sh.Cells.Item($r, 14).Value2 = $usd
                        $found = $true
                        break
                    }
                }
            }
            if ($found) { break }
        }
        if ($found) { break }
    }
    
    if ($found) {
        $wb.Save()
        Write-Output "DONE"
    }

} catch {
    # Error
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

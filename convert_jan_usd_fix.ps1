
$ErrorActionPreference = "Stop"

try {
    $filePattern = "*01*Report*.xlsx" # 01월 경영 지표 포함 파일
    $files = Get-ChildItem -Path "C:\Users\307984\.openclaw\workspace\Working" -Recurse -Filter $filePattern
    if ($files.Count -eq 0) { throw "File not found" }
    $targetFile = $files[0].FullName
    
    $newFile = "C:\Users\307984\.openclaw\workspace\Working\20260306_01월말 재고 꼬리표 Rpeort_USD.xlsx"
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $wb = $excel.Workbooks.Open($targetFile)
    $sh = $wb.Sheets.Item(1)
    
    # Header found at R13 C3
    $startRow = 14
    $c = 3 # Col C (회계재고가)
    
    $r = $startRow
    while ($true) {
        $name = $sh.Cells.Item($r, 2).Value2
        if (-not $name) { break }
        
        $valKrw100M = $sh.Cells.Item($r, $c).Value2
        
        # Check if value is numeric
        if ($valKrw100M -is [System.ValueType] -or $valKrw100M -match "^[\d\.]+$") {
            # 100M KRW -> KRW
            $valKrw = $valKrw100M * 100000000
            
            # KRW -> USD (Rate: 1427)
            $valUsd = $valKrw / 1427
            
            # USD -> M$
            $valUsdM = $valUsd / 1000000
            
            $sh.Cells.Item($r, $c).Value2 = $valUsdM
            Write-Host "Row ${r}: $valKrw100M (100M KRW) -> $valUsdM (M USD)"
        }
        
        $r++
    }
    
    $wb.SaveAs($newFile)
    Write-Host "File Saved: $newFile"
    
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
} catch {
    Write-Error $_
}

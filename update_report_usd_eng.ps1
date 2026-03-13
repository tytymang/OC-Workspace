
$ErrorActionPreference = "Stop"

# Use try-catch for error handling
try {
    # 1. Find File
    $files = Get-ChildItem -Path "C:\Users\307984\.openclaw\workspace\Working" -Recurse -Filter "*20260305_02*.xlsx"
    if ($files.Count -eq 0) {
        throw "Files not found."
    }
    $targetFile = $files[0].FullName
    Write-Host "Target: $targetFile"

    # 2. Open Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($targetFile)
    $sheet = $workbook.Sheets.Item(1)
    Write-Host "Sheet: $($sheet.Name)"

    # 3. Search Data (Unicode)
    $searchName = [char]0xC120 + [char]0xBD09 + [char]0xADDC
    $searchHeader = [char]0xD68C + [char]0xACC4 + [char]0xC7AC + [char]0xACE0 + [char]0xAC00
    
    $headerRow = 0
    $headerCol = 0
    
    for ($r=1; $r -le 10; $r++) {
        for ($c=1; $c -le 20; $c++) {
            $val = $sheet.Cells.Item($r, $c).Value2
            if ($val -ne $null -and $val.ToString().Contains($searchHeader)) {
                $headerRow = $r
                $headerCol = $c
                Write-Host "Header found at R$($r) C$($c)"
                break
            }
        }
        if ($headerRow -gt 0) { break }
    }
    
    if ($headerRow -eq 0) { throw "Header not found" }
    
    $nameRow = 0
    $nameCol = 0
    
    for ($r=$headerRow + 1; $r -le 100; $r++) {
        for ($c=1; $c -le 5; $c++) {
            $val = $sheet.Cells.Item($r, $c).Value2
            if ($val -eq $searchName) {
                $nameRow = $r
                $nameCol = $c
                Write-Host "Name found at R$($r) C$($c)"
                break
            }
        }
        if ($nameRow -gt 0) { break }
    }
    
    if ($nameRow -eq 0) { throw "Name not found" }

    # 4. Read Value
    $targetCell = $sheet.Cells.Item($nameRow, $headerCol)
    $currentValue = $targetCell.Value2
    Write-Host "Current Value: $currentValue"

    if ($currentValue -eq $null) {
        Write-Host "Value is null."
    } else {
        # 5. Calculate and Update
        $exchangeRate = 1427
        $krwValue = $currentValue * 100000000
        $usdValue = $krwValue / $exchangeRate
        
        Write-Host "Calculated USD: $usdValue"
        
        $targetCell.Value2 = $usdValue
        $workbook.Save()
        Write-Host "Saved."
    }

} catch {
    Write-Error $_
} finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { 
        $excel.Quit() 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

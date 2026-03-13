
$ErrorActionPreference = "Stop"

try {
    # 1. ?뚯씪 李얘린
    $files = Get-ChildItem -Path "C:\Users\307984\.openclaw\workspace\Working" -Recurse -Filter "*20260305_02*.xlsx"
    if ($files.Count -eq 0) {
        Write-Error "Files not found."
    }
    $targetFile = $files[0].FullName
    Write-Host "Target File: $targetFile"

    # 2. ?묒? ?닿린
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($targetFile)
    $sheet = $workbook.Sheets.Item(1)
    Write-Host "Sheet Name: $($sheet.Name)"

    # 3. ?곗씠??寃??(?좊땲肄붾뱶 ?ъ슜)
    # [char]0xC120 + [char]0xBD09 + [char]0xADDC # ?좊큺洹?    # [char]0xD68C + [char]0xACC4 + [char]0xC7AC + [char]0xACE0 + [char]0xAC00 # ?뚭퀎?ш퀬媛
    $searchName = [char]0xC120 + [char]0xBD09 + [char]0xADDC
    $searchHeader = [char]0xD68C + [char]0xACC4 + [char]0xC7AC + [char]0xACE0 + [char]0xAC00
    
    # ?ㅻ뜑 李얘린 (1~10??寃??
    $headerRow = 0
    $headerCol = 0
    
    for ($r=1; $r -le 10; $r++) {
        for ($c=1; $c -le 20; $c++) {
            $val = $sheet.Cells.Item($r, $c).Value2
            if ($val -ne $null -and $val.ToString().Contains($searchHeader)) {
                $headerRow = $r
                $headerCol = $c
                Write-Host "Potential Header found at R$($headerRow) C$($headerCol): $val"
                break
            }
        }
        if ($headerRow -gt 0) { break }
    }
    
    if ($headerRow -eq 0) { Write-Error "Header ($searchHeader) not found" }
    
    # ?대쫫 李얘린
    $nameRow = 0
    $nameCol = 0
    
    # ?ㅻ뜑 ???ㅼ쓬遺??寃??    for ($r=$headerRow + 1; $r -le 100; $r++) {
        for ($c=1; $c -le 5; $c++) {
            $val = $sheet.Cells.Item($r, $c).Value2
            if ($val -eq $searchName) {
                $nameRow = $r
                $nameCol = $c
                Write-Host "Name found at R$($nameRow) C$($nameCol)"
                break
            }
        }
        if ($nameRow -gt 0) { break }
    }
    
    if ($nameRow -eq 0) { Write-Error "Name ($searchName) not found" }

    # 4. 媛??쎄린
    $targetCell = $sheet.Cells.Item($nameRow, $headerCol)
    $currentValue = $targetCell.Value2
    Write-Host "Current Value at R$($nameRow) C$($headerCol): $currentValue"

    if ($currentValue -eq $null) {
        Write-Host "Value is null. Checking nearby cells..."
        $valRight = $sheet.Cells.Item($nameRow, $headerCol + 1).Value2
        $valLeft = $sheet.Cells.Item($nameRow, $headerCol - 1).Value2
        Write-Host "Right: $valRight, Left: $valLeft"
    }

    # ?섏쑉 ?곸슜 (1427??
    if ($currentValue -ne $null) {
        $exchangeRate = 1427
        $krwValue = $currentValue * 100000000 # ?듭썝 -> ??        $usdValue = $krwValue / $exchangeRate
        
        Write-Host "Calculated USD: $usdValue"
        
        # 5. 媛??낅뜲?댄듃
        $targetCell.Value2 = $usdValue
        $workbook.Save()
        Write-Host "File Saved."
    }

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { 
        $excel.Quit() 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}


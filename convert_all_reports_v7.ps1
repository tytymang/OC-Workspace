
$ErrorActionPreference = "Stop"

function Convert-Report {
    param (
        [string]$SourceFile,
        [string]$DestFile,
        [double]$Rate
    )

    Write-Host "Processing $SourceFile"
    Write-Host "Rate: $Rate"

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $wb = $excel.Workbooks.Open($SourceFile)
        $sh = $wb.Sheets.Item(1)

        $usedRange = $sh.UsedRange
        
        $headerRow = 0
        $amountCols = @()
        
        # 1. Header Search (Only "억원" -> "M$")
        # Unicode for '억원' = C5B5 + C6D0
        $targetStr = [char]0xC5B5 + [char]0xC6D0
        
        for ($r = 1; $r -le 20; $r++) {
            for ($c = 1; $c -le 20; $c++) {
                $cell = $sh.Cells.Item($r, $c)
                $val = $cell.Value2
                
                if ($val -is [string]) {
                     if ($val.Contains($targetStr)) {
                        $newHeader = $val.Replace($targetStr, "M$")
                        $cell.Value2 = $newHeader
                        $amountCols += $c
                        $headerRow = $r # Keep last found header row as reference
                        Write-Host "Updated Header at R${r}C${c}: $val -> $newHeader"
                     }
                }
            }
        }
        
        # 2. Convert Data
        # Filter amountCols to exclude description columns (Col 11 seems to be description based on error log)
        # Column 11 contained "FAB 1 GT..." text, which caused conversion error
        
        # Heuristic: Check Row 14 (Sample Data Row) to see if it's numeric
        $validCols = @()
        if ($headerRow -gt 0) {
            $sampleRow = $headerRow + 1
            foreach ($c in $amountCols) {
                $sampleVal = $sh.Cells.Item($sampleRow, $c).Value2
                if ($sampleVal -is [System.ValueType]) { # Numeric
                    $validCols += $c
                } else {
                    Write-Warning "Skipping Col $c (Sample value is not numeric: $sampleVal)"
                }
            }
        }
        
        if ($validCols.Count -gt 0) {
            $amountCols = $validCols | Select-Object -Unique
            $startRow = $headerRow + 1
            
            # Find last row
            $lastRow = $startRow
            while ($true) {
                $name = $sh.Cells.Item($lastRow, 2).Value2
                if ([string]::IsNullOrWhiteSpace($name)) { break }
                $lastRow++
            }
            $lastRow-- 
            
            Write-Host "Converting Rows $startRow to $lastRow in Cols $($amountCols -join ',')"
            
            for ($r = $startRow; $r -le $lastRow; $r++) {
                foreach ($c in $amountCols) {
                    $cell = $sh.Cells.Item($r, $c)
                    $val = $cell.Value2
                    
                    if ($val -is [System.ValueType]) { # Numeric check
                        # KRW(100M) -> KRW -> USD -> M USD
                        $newVal = ($val * 100000000) / $Rate / 1000000
                        $cell.Value2 = $newVal
                    }
                }
            }
        }
        
        # Save
        $wb.SaveAs($DestFile)
        Write-Host "Saved: $DestFile"
        
        $wb.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        
    } catch {
        Write-Error "Error: $_"
        if ($wb) { $wb.Close($false) }
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

$janRate = 1427
$febRate = 1424.5

$files = Get-ChildItem "C:\Users\307984\.openclaw\workspace\Working" -Recurse

# Jan
$janFiles = $files | Where-Object { $_.Name -like "*01*Report*.xlsx" -and $_.Name -notlike "*USD*" }
if ($janFiles) {
    $janSource = $janFiles[0].FullName
    $janDest = "C:\Users\307984\.openclaw\workspace\Working\20260306_01월말 재고 꼬리표 Report_USD.xlsx"
    if (Test-Path $janDest) { Remove-Item $janDest -Force }
    Convert-Report -SourceFile $janSource -DestFile $janDest -Rate $janRate
}

# Feb
$febFiles = $files | Where-Object { $_.Name -like "*02*Report*.xlsx" -and $_.Name -notlike "*USD*" }
if ($febFiles) {
    $febSource = $febFiles[0].FullName
    $febDest = "C:\Users\307984\.openclaw\workspace\Working\20260306_02월말 재고 꼬리표 Report_USD.xlsx"
    if (Test-Path $febDest) { Remove-Item $febDest -Force }
    Convert-Report -SourceFile $febSource -DestFile $febDest -Rate $febRate
}

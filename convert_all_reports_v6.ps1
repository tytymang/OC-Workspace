
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
        
        # 1. Header Search
        # Match '억원' -> 'M$'
        for ($r = 1; $r -le 20; $r++) {
            for ($c = 1; $c -le 20; $c++) {
                $cell = $sh.Cells.Item($r, $c)
                $val = $cell.Value2
                
                if ($val -ne $null -and $val -is [string]) {
                     # Unicode-safe match for '억원'
                     $sVal = $val.ToString()
                     $targetStr = [char]0xC5B5 + [char]0xC6D0 # 억원
                     
                     if ($sVal.Contains($targetStr)) {
                        $newHeader = $sVal.Replace($targetStr, "M$")
                        $cell.Value2 = $newHeader
                        
                        # Only mark as amount column if header implies it's a value column
                        # (e.g. not a description column)
                        # Heuristic: Check next row value. If number, it's an amount column.
                        $nextVal = $sh.Cells.Item($r + 1, $c).Value2
                        if ($nextVal -is [System.ValueType] -or $nextVal -eq $null) {
                            $amountCols += $c
                            $headerRow = $r
                            Write-Host "Updated Header (Amount Col) at R${r}C${c}: $sVal -> $newHeader"
                        } else {
                            Write-Host "Updated Header (Text Col?) at R${r}C${c}: $sVal -> $newHeader"
                        }
                     }
                }
            }
        }
        
        # 2. Convert Data
        if ($headerRow -gt 0) {
            $amountCols = $amountCols | Select-Object -Unique
            $startRow = $headerRow + 1
            
            # Find last row
            $lastRow = $startRow
            while ($true) {
                # Check name column (Col 2)
                $name = $sh.Cells.Item($lastRow, 2).Value2
                
                # Stop if empty name
                if ([string]::IsNullOrWhiteSpace($name)) { break }
                $lastRow++
            }
            $lastRow-- 
            
            if ($lastRow -ge $startRow) {
                Write-Host "Converting Rows $startRow to $lastRow in Cols $($amountCols -join ',')"
                
                for ($r = $startRow; $r -le $lastRow; $r++) {
                    foreach ($c in $amountCols) {
                        $cell = $sh.Cells.Item($r, $c)
                        $val = $cell.Value2
                        
                        # Only convert numeric values
                        if ($val -is [System.ValueType]) { # Double, Int, etc.
                            # KRW(100M) -> KRW -> USD -> M USD
                            $newVal = ([double]$val * 100000000) / $Rate / 1000000
                            $cell.Value2 = $newVal
                        }
                    }
                }
            }
        } else {
            Write-Warning "No header found for conversion."
        }
        
        # Save logic
        $dir = [System.IO.Path]::GetDirectoryName($DestFile)
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
        
        $tempFile = [System.IO.Path]::Combine($dir, "temp_" + [System.Guid]::NewGuid().ToString() + ".xlsx")
        $wb.SaveAs($tempFile)
        $wb.Close($false)
        
        if (Test-Path $DestFile) { Remove-Item $DestFile -Force }
        Move-Item $tempFile $DestFile -Force
        Write-Host "Saved: $DestFile"

    } catch {
        Write-Error "Error: $_"
        if ($wb) { $wb.Close($false) }
    } finally {
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
    Convert-Report -SourceFile $janSource -DestFile $janDest -Rate $janRate
}

# Feb
$febFiles = $files | Where-Object { $_.Name -like "*02*Report*.xlsx" -and $_.Name -notlike "*USD*" }
if ($febFiles) {
    $febSource = $febFiles[0].FullName
    $febDest = "C:\Users\307984\.openclaw\workspace\Working\20260306_02월말 재고 꼬리표 Report_USD.xlsx"
    Convert-Report -SourceFile $febSource -DestFile $febDest -Rate $febRate
}

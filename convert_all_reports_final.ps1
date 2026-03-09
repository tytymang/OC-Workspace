
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
        $targetStr = [char]0xC5B5 + [char]0xC6D0 # 억원
        
        for ($r = 1; $r -le 20; $r++) {
            for ($c = 1; $c -le 20; $c++) {
                $cell = $sh.Cells.Item($r, $c)
                $val = $cell.Value2
                
                if ($val -ne $null -and $val -is [string]) {
                     if ($val.Contains($targetStr)) {
                        $newHeader = $val.Replace($targetStr, "M$")
                        $cell.Value2 = $newHeader
                        $amountCols += $c
                        $headerRow = $r
                        Write-Host "Header: R${r}C${c}"
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
                $name = $sh.Cells.Item($lastRow, 2).Value2
                if ([string]::IsNullOrWhiteSpace($name)) { break }
                $lastRow++
            }
            $lastRow-- 
            
            # Filter non-numeric columns
            $validCols = @()
            foreach ($c in $amountCols) {
                # Check first data row
                $sampleVal = $sh.Cells.Item($startRow, $c).Value2
                if ($sampleVal -is [System.ValueType] -or ($sampleVal -is [string] -and $sampleVal -match "^[\d\.]+$")) {
                    $validCols += $c
                } else {
                    Write-Warning "Skipping Col $c (Sample: $sampleVal)"
                }
            }
            
            if ($validCols.Count -gt 0) {
                Write-Host "Converting Rows $startRow to $lastRow in Cols $($validCols -join ',')"
                
                for ($r = $startRow; $r -le $lastRow; $r++) {
                    foreach ($c in $validCols) {
                        $cell = $sh.Cells.Item($r, $c)
                        $val = $cell.Value2
                        
                        if ($val -is [System.ValueType]) {
                            # Explicit cast to double
                            $dVal = [double]$val
                            $newVal = ($dVal * 100000000.0) / $Rate / 1000000.0
                            $cell.Value2 = $newVal
                        }
                    }
                }
            }
        } else {
            Write-Warning "No header found."
        }
        
        # Save (as new file)
        $wb.SaveAs($DestFile)
        Write-Host "Saved: $DestFile"
        
        $wb.Close($false)
    } catch {
        Write-Error "Error: $_"
        if ($wb) { $wb.Close($false) }
    } finally {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

$janRate = 1427.0
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

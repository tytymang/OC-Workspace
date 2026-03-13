
$ErrorActionPreference = "Stop"

function Convert-Report {
    param (
        [string]$SourceFile,
        [string]$DestFile,
        [double]$Rate,
        [string]$MonthName
    )

    Write-Host "Processing $MonthName Report..."
    Write-Host "Source: $SourceFile"
    Write-Host "Dest: $DestFile"
    Write-Host "Rate: $Rate"

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $wb = $excel.Workbooks.Open($SourceFile)
        $sh = $wb.Sheets.Item(1) # Assuming first sheet is the target "보고서"

        # 1. Update Headers (억원 -> M$)
        # Scan first 20 rows, 20 cols
        $usedRange = $sh.UsedRange
        $maxRow = $usedRange.Rows.Count
        $maxCol = $usedRange.Columns.Count
        if ($maxRow -gt 100) { $maxRow = 100 } # Limit scan for headers

        # Find header row and amount columns
        $headerRow = 0
        $amountCols = @()

        for ($r = 1; $r -le 20; $r++) {
            for ($c = 1; $c -le 20; $c++) {
                $cell = $sh.Cells.Item($r, $c)
                $val = $cell.Value2
                
                if ($val -is [string]) {
                    if ($val -match "억원") {
                        # Change header text
                        $newHeader = $val -replace "억원", "M$"
                        $cell.Value2 = $newHeader
                        Write-Host "  Updated Header at [$r,$c]: $val -> $newHeader"
                        
                        # Mark this column as an amount column
                        $amountCols += $c
                        $headerRow = $r
                    }
                    elseif ($val -match "회계재고가" -or $val -match "재고가" -or $val -match "합의" -or $val -match "금액") {
                        # If unit is not in header but it implies amount, check next row or unit row
                        # For now, rely on "억원" being in the header based on user prompt
                        # "필드명에 '억원' 이라고 되어 있는 것을 'M$' 로 변경"
                    }
                }
            }
        }

        if ($amountCols.Count -eq 0) {
            Write-Warning "No headers with '억원' found. Checking for '회계재고가' etc."
            # Fallback: Find specific column names and assume they need conversion
             for ($r = 1; $r -le 20; $r++) {
                for ($c = 1; $c -le 20; $c++) {
                    $val = $sh.Cells.Item($r, $c).Value2
                    if ($val -is [string] -and ($val -match "재고가" -or $val -match "금액")) {
                         $amountCols += $c
                         $headerRow = $r
                         Write-Host "  Found Amount Column at [$r,$c]: $val"
                         # Append (M$) if not present
                         if ($val -notmatch "M\$") {
                             $sh.Cells.Item($r, $c).Value2 = "$val (M$)"
                         }
                    }
                }
             }
        }

        # 2. Convert Data
        if ($headerRow -gt 0 -and $amountCols.Count -gt 0) {
            $startRow = $headerRow + 1
            $lastRow = $sh.UsedRange.Rows.Count
            
            # Remove duplicate cols
            $amountCols = $amountCols | Select-Object -Unique

            Write-Host "  Converting data from Row $startRow to $lastRow in cols: $($amountCols -join ', ')"

            for ($r = $startRow; $r -le $lastRow; $r++) {
                # Check if row is valid (has a name in col 2 or 1)
                $name = $sh.Cells.Item($r, 2).Value2
                if ([string]::IsNullOrWhiteSpace($name)) { continue }

                foreach ($c in $amountCols) {
                    $cell = $sh.Cells.Item($r, $c)
                    $val = $cell.Value2
                    
                    if ($val -is [System.ValueType]) { # Numeric
                        # KRW (100M) -> KRW -> USD -> M USD
                        # val * 100,000,000 / Rate / 1,000,000
                        # = val * 100 / Rate
                        
                        $newVal = ($val * 100000000) / $Rate / 1000000
                        $cell.Value2 = $newVal
                    }
                }
            }
        } else {
            Write-Warning "Could not identify data columns."
        }

        # 3. Save
        $wb.SaveAs($DestFile)
        Write-Host "Saved: $DestFile"

    } catch {
        Write-Error "Error processing $MonthName report: $_"
    } finally {
        if ($wb) { $wb.Close($false) }
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

# Parameters
$janRate = 1427
$febRate = 1424.5

# Find Files
$files = Get-ChildItem "C:\Users\307984\.openclaw\workspace\Working" -Recurse

# Jan File
$janSource = ($files | Where-Object { $_.Name -like "*01*Report*.xlsx" -and $_.Name -notlike "*USD*" })[0].FullName
$janDest = "C:\Users\307984\.openclaw\workspace\Working\20260306_01월말 재고 꼬리표 Report_USD.xlsx"

# Feb File
$febSource = ($files | Where-Object { $_.Name -like "*02*Report*.xlsx" -and $_.Name -notlike "*USD*" })[0].FullName
$febDest = "C:\Users\307984\.openclaw\workspace\Working\20260306_02월말 재고 꼬리표 Report_USD.xlsx"

# Execute
if ($janSource) {
    Convert-Report -SourceFile $janSource -DestFile $janDest -Rate $janRate -MonthName "January"
} else {
    Write-Error "January Source File not found."
}

if ($febSource) {
    Convert-Report -SourceFile $febSource -DestFile $febDest -Rate $febRate -MonthName "February"
} else {
    Write-Error "February Source File not found."
}

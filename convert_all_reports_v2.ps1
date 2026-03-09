
$ErrorActionPreference = "Stop"

function Convert-Report {
    param (
        [string]$SourceFile,
        [string]$DestFile,
        [double]$Rate,
        [string]$MonthName
    )

    Write-Host "Processing $MonthName Report..."
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $wb = $excel.Workbooks.Open($SourceFile)
        $sh = $wb.Sheets.Item(1)

        $usedRange = $sh.UsedRange
        $maxRow = 20
        $maxCol = 20
        
        $headerRow = 0
        $amountCols = @()
        
        # 1. Header Search
        # Match '억원' or '회계재고가' etc.
        for ($r = 1; $r -le $maxRow; $r++) {
            for ($c = 1; $c -le $maxCol; $c++) {
                $cell = $sh.Cells.Item($r, $c)
                $val = $cell.Value2
                
                if ($val -is [string]) {
                    if ($val -match "억원") {
                        $newHeader = $val -replace "억원", "M$"
                        $cell.Value2 = $newHeader
                        $amountCols += $c
                        $headerRow = $r
                    }
                    elseif ($val -match "회계재고가" -or $val -match "재고가" -or $val -match "합의" -or $val -match "금액") {
                        # If unit is implied but not explicit, assume convert
                        # But user said "필드명에 '억원' 이라고 되어 있는 것을 'M$' 로 변경"
                        # So prioritize '억원' match. If none found, fallback.
                    }
                }
            }
        }
        
        # Fallback if no explicit '억원' found
        if ($amountCols.Count -eq 0) {
            Write-Host "No explicit '억원' headers. Searching for amount columns..."
            for ($r = 1; $r -le $maxRow; $r++) {
                for ($c = 1; $c -le $maxCol; $c++) {
                    $cell = $sh.Cells.Item($r, $c)
                    $val = $cell.Value2
                    if ($val -is [string] -and ($val -match "재고가" -or $val -match "금액")) {
                         $amountCols += $c
                         $headerRow = $r
                         # Append (M$)
                         if ($val -notmatch "M\$") {
                             $cell.Value2 = "$val (M$)"
                         }
                    }
                }
            }
        }
        
        # 2. Convert Data
        if ($amountCols.Count -gt 0) {
            $amountCols = $amountCols | Select-Object -Unique
            $startRow = $headerRow + 1
            $lastRow = $sh.UsedRange.Rows.Count
            if ($lastRow -gt 100) { $lastRow = 100 } # Safety limit for now
            
            Write-Host "Converting Rows $startRow to $lastRow in Cols $($amountCols -join ',')"
            
            for ($r = $startRow; $r -le $lastRow; $r++) {
                # Check for valid name in col 2
                $name = $sh.Cells.Item($r, 2).Value2
                if (-not $name) { continue }
                
                foreach ($c in $amountCols) {
                    $cell = $sh.Cells.Item($r, $c)
                    $val = $cell.Value2
                    
                    if ($val -is [System.ValueType] -and $val -gt 0) {
                        # KRW(100M) -> KRW -> USD -> M USD
                        $newVal = ($val * 100000000) / $Rate / 1000000
                        $cell.Value2 = $newVal
                    }
                }
            }
        }

        $wb.SaveAs($DestFile)
        Write-Host "Saved: $DestFile"

    } catch {
        Write-Error "Error in $MonthName: $_"
    } finally {
        if ($wb) { $wb.Close($false) }
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

$janRate = 1427
$febRate = 1424.5

$files = Get-ChildItem "C:\Users\307984\.openclaw\workspace\Working" -Recurse

# Jan Source
$janSource = ($files | Where-Object { $_.Name -like "*01*Report*.xlsx" -and $_.Name -notlike "*USD*" })[0].FullName
$janDest = "C:\Users\307984\.openclaw\workspace\Working\20260306_01월말 재고 꼬리표 Report_USD.xlsx"

if ($janSource) {
    Convert-Report -SourceFile $janSource -DestFile $janDest -Rate $janRate -MonthName "January"
}

# Feb Source
$febSource = ($files | Where-Object { $_.Name -like "*02*Report*.xlsx" -and $_.Name -notlike "*USD*" })[0].FullName
$febDest = "C:\Users\307984\.openclaw\workspace\Working\20260306_02월말 재고 꼬리표 Report_USD.xlsx"

if ($febSource) {
    Convert-Report -SourceFile $febSource -DestFile $febDest -Rate $febRate -MonthName "February"
}

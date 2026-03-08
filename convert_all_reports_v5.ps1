
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
        
        # 1. Header Search (Simple string match)
        for ($r = 1; $r -le 20; $r++) {
            for ($c = 1; $c -le 20; $c++) {
                $cell = $sh.Cells.Item($r, $c)
                $val = $cell.Value2
                
                if ($val -ne $null) {
                     # Unicode-safe match
                     if ($val.ToString().Contains([char]0xC5B5 + [char]0xC6D0)) { # 억원
                        $newHeader = $val.ToString().Replace(([char]0xC5B5 + [char]0xC6D0), "M$")
                        $cell.Value2 = $newHeader
                        $amountCols += $c
                        $headerRow = $r
                        Write-Host "Updated Header at R${r}C${c}: $val -> $newHeader"
                     }
                }
            }
        }
        
        # Fallback
        if ($amountCols.Count -eq 0) {
            Write-Host "No explicit '억원' headers. Searching for amount columns..."
            for ($r = 1; $r -le 20; $r++) {
                for ($c = 1; $c -le 20; $c++) {
                    $cell = $sh.Cells.Item($r, $c)
                    $val = $cell.Value2
                    
                    if ($val -ne $null) {
                        # 재고가: [char]0xC7AC + [char]0xACE0 + [char]0xAC00
                        # 금액: [char]0xAE08 + [char]0xC561
                        
                        $sVal = $val.ToString()
                        if ($sVal.Contains([char]0xC7AC + [char]0xACE0 + [char]0xAC00) -or $sVal.Contains([char]0xAE08 + [char]0xC561)) {
                             $amountCols += $c
                             $headerRow = $r
                             if (-not $sVal.Contains("M$")) {
                                 $cell.Value2 = "$val (M$)"
                                 Write-Host "Updated Header at R${r}C${c}: $val -> $val (M$)"
                             }
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
                $name = $sh.Cells.Item($lastRow, 2).Value2
                if (-not $name) { break }
                $lastRow++
            }
            $lastRow-- 
            
            Write-Host "Converting Rows $startRow to $lastRow in Cols $($amountCols -join ',')"
            
            for ($r = $startRow; $r -le $lastRow; $r++) {
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
        } else {
            Write-Warning "No header found for conversion."
        }
        
        # Force safe filename
        # Ensure directory exists
        $dir = [System.IO.Path]::GetDirectoryName($DestFile)
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
        
        # Save as temp English name first
        $tempFile = [System.IO.Path]::Combine($dir, "temp_report_usd.xlsx")
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force }
        
        $wb.SaveAs($tempFile)
        $wb.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        
        # Rename to target file
        if (Test-Path $DestFile) { Remove-Item $DestFile -Force }
        Move-Item $tempFile $DestFile -Force
        Write-Host "Saved: $DestFile"

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

# Jan Source
$janFiles = $files | Where-Object { $_.Name -like "*01*Report*.xlsx" -and $_.Name -notlike "*USD*" }
if ($janFiles) {
    $janSource = $janFiles[0].FullName
    # Correct filename for Jan
    $janDest = "C:\Users\307984\.openclaw\workspace\Working\20260306_01월말 재고 꼬리표 Report_USD.xlsx"
    Convert-Report -SourceFile $janSource -DestFile $janDest -Rate $janRate
}

# Feb Source
$febFiles = $files | Where-Object { $_.Name -like "*02*Report*.xlsx" -and $_.Name -notlike "*USD*" }
if ($febFiles) {
    $febSource = $febFiles[0].FullName
    # Correct filename for Feb
    $febDest = "C:\Users\307984\.openclaw\workspace\Working\20260306_02월말 재고 꼬리표 Report_USD.xlsx"
    Convert-Report -SourceFile $febSource -DestFile $febDest -Rate $febRate
}

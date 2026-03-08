
$ErrorActionPreference = "Stop"

function Convert-One-File {
    param (
        [string]$Path,
        [string]$Dest,
        [double]$Rate
    )
    
    Write-Host "Processing: $Path"
    Write-Host "Rate: $Rate"

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    try {
        $wb = $excel.Workbooks.Open($Path)
        $sh = $wb.Sheets.Item(1)
        
        # 1. Headers (Only search for 억원 headers)
        $headerRow = 0
        $cols = @()
        
        $targetStr = [char]0xC5B5 + [char]0xC6D0 # 억원
        
        for ($r=1; $r -le 20; $r++) {
            for ($c=1; $c -le 20; $c++) {
                $cell = $sh.Cells.Item($r, $c)
                $v = $cell.Value2
                
                if ($v -ne $null -and $v -is [string]) {
                    if ($v.Contains($targetStr)) {
                        $newH = $v.Replace($targetStr, "M$")
                        $cell.Value2 = $newH
                        $cols += $c
                        $headerRow = $r
                    }
                }
            }
        }
        
        # 2. Convert
        if ($headerRow -gt 0) {
            $startRow = $headerRow + 1
            $lastRow = $startRow
            while ($true) {
                $n = $sh.Cells.Item($lastRow, 2).Value2
                if ([string]::IsNullOrWhiteSpace($n)) { break }
                $lastRow++
            }
            $lastRow--
            
            if ($cols.Count -gt 0) {
                Write-Host "Converting Rows $startRow to $lastRow"
                
                for ($r = $startRow; $r -le $lastRow; $r++) {
                    foreach ($c in $cols) {
                        $cell = $sh.Cells.Item($r, $c)
                        $val = $cell.Value2
                        
                        # Only convert numeric
                        if ($val -ne $null) {
                            # Attempt cast
                            try {
                                # Use decimal/double cast
                                $dVal = [double]$val
                                # (Val * 100M) / Rate / 1M = Val * 100 / Rate
                                $newVal = ($dVal * 100.0) / $Rate
                                $cell.Value2 = $newVal
                            } catch {
                                # Skip non-numerics
                            }
                        }
                    }
                }
            }
        }
        
        $wb.SaveAs($Dest)
        Write-Host "Saved to $Dest"
        $wb.Close($false)
        
    } catch {
        Write-Error $_
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
$jFiles = $files | Where-Object { $_.Name -like "*01*Report*.xlsx" -and $_.Name -notlike "*USD*" }
if ($jFiles) {
    $src = $jFiles[0].FullName
    $dst = "C:\Users\307984\.openclaw\workspace\Working\20260306_01월말 재고 꼬리표 Report_USD.xlsx"
    if (Test-Path $dst) { Remove-Item $dst -Force }
    Convert-One-File -Path $src -Dest $dst -Rate $janRate
}

# Feb
$fFiles = $files | Where-Object { $_.Name -like "*02*Report*.xlsx" -and $_.Name -notlike "*USD*" }
if ($fFiles) {
    $src = $fFiles[0].FullName
    $dst = "C:\Users\307984\.openclaw\workspace\Working\20260306_02월말 재고 꼬리표 Report_USD.xlsx"
    if (Test-Path $dst) { Remove-Item $dst -Force }
    Convert-One-File -Path $src -Dest $dst -Rate $febRate
}

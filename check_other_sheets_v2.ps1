
$ErrorActionPreference = "Stop"

try {
    # 1. 환율 파일 찾기
    $files = Get-ChildItem -Path "C:\Users\307984\.openclaw\workspace\Working" -Recurse -Include *202602*.xlsx
    $rateFile = ($files | Where-Object { $_.Name -like "*ȯ*" -or $_.Name -like "*exchange*" -or $_.Name -like "*rate*" })[0].FullName
    if (-not $rateFile) { $rateFile = $files[-1].FullName }
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $wb = $excel.Workbooks.Open($rateFile)
    
    # 2. 다른 시트 확인 (Sheet2, Sheet3 등)
    foreach ($sh in $wb.Sheets) {
        Write-Host "Checking Sheet: $($sh.Name)"
        $vals = $sh.Range("A1:K20").Value2
        
        # Check for 2025 or 2026 in first few rows
        if ($vals -is [System.Array] -and $vals.Rank -eq 2) {
            for ($r=1; $r -le $vals.GetLength(0); $r++) {
                $rowStr = ""
                for ($c=1; $c -le $vals.GetLength(1); $c++) {
                    $v = $vals[$r,$c]
                    if ($v) { $rowStr += "$v | " }
                }
                if ($rowStr -match "2025" -or $rowStr -match "2026") {
                    Write-Host "Found year in $($sh.Name) Row ${r}: $rowStr"
                    
                    # Dump this sheet's first 20 rows
                    Write-Host "Dumping $($sh.Name)..."
                    for ($dr=1; $dr -le 20; $dr++) {
                        $dStr = ""
                        for ($dc=1; $dc -le 10; $dc++) {
                             $dv = $sh.Cells.Item($dr, $dc).Value2
                             if ($dv) { $dStr += "[$dr,$dc]: $dv | " }
                        }
                        if ($dStr) { Write-Host $dStr }
                    }
                }
            }
        }
    }
    
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
} catch {
    Write-Error $_
}

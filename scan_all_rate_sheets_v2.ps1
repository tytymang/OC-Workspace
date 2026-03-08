
$ErrorActionPreference = "Stop"

try {
    # 1. Find File
    $files = Get-ChildItem "C:\Users\307984\.openclaw\workspace\Working" -Recurse -Filter "*환율*.xlsx"
    if ($files.Count -eq 0) { throw "File not found" }
    $rateFile = $files[0].FullName
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $wb = $excel.Workbooks.Open($rateFile)
    
    Write-Host "File: $rateFile"
    
    foreach ($sh in $wb.Sheets) {
        Write-Host "Sheet: $($sh.Name)"
        
        $vals = $sh.Range("A1:K20").Value2
        if ($vals) {
           if ($vals.Rank -eq 2) {
               for ($r=1; $r -le $vals.GetLength(0); $r++) {
                    $rowStr = ""
                    for ($c=1; $c -le $vals.GetLength(1); $c++) {
                        $v = $vals[$r, $c]
                        if ($v) { $rowStr += "[$r,$c]: $v | " }
                    }
                    if ($rowStr -match "2026") {
                        Write-Host "  Found 2026 in Row ${r}: $rowStr"
                    }
                    if ($rowStr -match "1월" -or $rowStr -match "Jan" -or $rowStr -match "01월") {
                        Write-Host "  Found Jan in Row ${r}: $rowStr"
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

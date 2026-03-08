
$ErrorActionPreference = "Stop"

try {
    # 1. 환율 파일 찾기
    $files = Get-ChildItem "C:\Users\307984\.openclaw\workspace\Working" -Recurse -Filter "*환율*.xlsx"
    if ($files.Count -eq 0) { throw "File not found" }
    $rateFile = $files[0].FullName
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $wb = $excel.Workbooks.Open($rateFile)
    
    # Check all sheets for "USD" and "2026"
    foreach ($sh in $wb.Sheets) {
        $val = $sh.UsedRange.Value2
        # Simple scan
        Write-Host "Sheet: $($sh.Name)"
        $sh.Range("A1:K20").Cells | ForEach-Object {
            $v = $_.Value2
            if ($v) {
                Write-Host "[$($_.Row),$($_.Column)]: $v"
            }
        }
    }
    
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
} catch {
    Write-Error $_
}


$ErrorActionPreference = "Stop"

try {
    # 1. 환율 파일 찾기
    $files = Get-ChildItem "C:\Users\307984\.openclaw\workspace\Working" -Recurse -Filter "*환율*.xlsx"
    if ($files.Count -eq 0) { throw "File not found" }
    $rateFile = $files[0].FullName
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    
    try {
        $wb = $excel.Workbooks.Open($rateFile)
        
        Write-Host "File: $rateFile"
        Write-Host "Sheets:"
        
        foreach ($sh in $wb.Sheets) {
            Write-Host "- $($sh.Name)"
            
            # 각 시트의 A1:K20 내용을 대략적으로 훑어 2026년 1월/2월 데이터 찾기
            $vals = $sh.Range("A1:K20").Value2
            
            if ($vals) {
                # 2D 배열인 경우 (대부분)
                if ($vals.Rank -eq 2) {
                   for ($r=1; $r -le $vals.GetLength(0); $r++) {
                        $rowStr = ""
                        for ($c=1; $c -le $vals.GetLength(1); $c++) {
                            $v = $vals[$r, $c]
                            if ($v) { $rowStr += "[$r,$c]: $v | " }
                        }
                        if ($rowStr -match "2026") {
                            Write-Host "  Found 2026 data in row $r: $rowStr"
                        }
                        if ($rowStr -match "1월" -or $rowStr -match "Jan") {
                            Write-Host "  Found Jan data in row $r: $rowStr"
                        }
                   }
                }
            }
        }
    } finally {
        if ($wb) { $wb.Close($false) }
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
} catch {
    Write-Error $_
}

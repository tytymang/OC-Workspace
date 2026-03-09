
$ErrorActionPreference = "Stop"

try {
    # 1. 환율 파일 찾기
    $rateFile = (Get-ChildItem "C:\Users\307984\.openclaw\workspace\Working" -Recurse -Filter "*환율*.xlsx")[0].FullName
    Write-Host "Exchange Rate File: $rateFile"

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wbRate = $excel.Workbooks.Open($rateFile)
    
    # 2. 환율 정보 읽기 (기말 환율)
    # 보통 환율표는 "USD" 행과 월별 열로 구성됨.
    # 파일 구조를 모르므로 일단 첫 시트의 내용을 덤프해서 위치 파악
    
    $shRate = $wbRate.Sheets.Item(1)
    Write-Host "Rate Sheet: $($shRate.Name)"
    
    # Dump first 20 rows & cols to find "USD" and months (2026.01, 2026.02 or similar)
    # And specifically "기말" (Ending) rate
    
    for ($r=1; $r -le 20; $r++) {
        $rowStr = ""
        for ($c=1; $c -le 20; $c++) {
            $val = $shRate.Cells.Item($r, $c).Value2
            if ($val) { $rowStr += "[$r,$c]: $val | " }
        }
        if ($rowStr) { Write-Host $rowStr }
    }
    
    $wbRate.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

} catch {
    Write-Error $_
}

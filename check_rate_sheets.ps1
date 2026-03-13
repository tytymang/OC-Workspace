
$ErrorActionPreference = "Stop"

try {
    # 1. 환율 파일 찾기
    $files = Get-ChildItem "C:\Users\307984\.openclaw\workspace\Working" -Recurse -Filter "*환율*.xlsx"
    $rateFile = $files[0].FullName
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $wb = $excel.Workbooks.Open($rateFile)
    
    # 2. 모든 시트 이름 확인
    Write-Host "Sheets in $rateFile"
    foreach ($sh in $wb.Sheets) {
        Write-Host "Sheet: $($sh.Name)"
    }
    
    # 3. '2026' 또는 '기말'이 포함된 시트나 데이터 찾기
    # 첫 번째 시트가 2019년 데이터라면 다른 시트에 2026년 데이터가 있을 수 있음.
    
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
} catch {
    Write-Error $_
}

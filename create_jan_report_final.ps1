
$ErrorActionPreference = "Stop"

try {
    # 1. 파일 경로 설정
    $targetFile = (Get-ChildItem -Path "C:\Users\307984\.openclaw\workspace\Working" -Recurse -Filter "*01*Report*.xlsx")[0].FullName
    
    # 영문 파일명으로 저장 (안전하게)
    $newFile = "C:\Users\307984\.openclaw\workspace\Working\20260306_Jan_Stock_Report_USD.xlsx"
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $wb = $excel.Workbooks.Open($targetFile)
    $sh = $wb.Sheets.Item(1)
    
    # 데이터 수정 (1월말 환율 1427원 적용)
    $r = 14
    while ($true) {
        $name = $sh.Cells.Item($r, 2).Value2
        if (-not $name) { break }
        
        $val = $sh.Cells.Item($r, 3).Value2
        if ($val -is [System.ValueType]) {
             # 억원 -> M$ (환율 1427)
             $newVal = ($val * 100000000) / 1427 / 1000000
             $sh.Cells.Item($r, 3).Value2 = $newVal
             Write-Host "Updated Row $r: $val -> $newVal"
        }
        $r++
    }
    
    $wb.SaveAs($newFile)
    Write-Host "Saved to: $newFile"
    
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    # 파일명 변경 (한글로)
    $koreanName = "C:\Users\307984\.openclaw\workspace\Working\20260306_01월말 재고 꼬리표 Report_USD.xlsx"
    Move-Item -Path $newFile -Destination $koreanName -Force
    Write-Host "Renamed to: $koreanName"

} catch {
    Write-Error $_
}


$ErrorActionPreference = "Stop"

try {
    $filePattern = "*01*Report*.xlsx"
    $files = Get-ChildItem -Path "C:\Users\307984\.openclaw\workspace\Working" -Recurse -Filter $filePattern
    if ($files.Count -eq 0) { throw "File not found" }
    $targetFile = $files[0].FullName
    
    # 수정: 파일명 오타 수정 (Rpeort -> Report) 및 경로 안전하게 설정
    $newFile = "C:\Users\307984\.openclaw\workspace\Working\20260306_01월말 재고 꼬리표 Report_USD.xlsx"
    
    # 덮어쓰기 방지를 위해 기존 파일 있으면 삭제
    if (Test-Path $newFile) { Remove-Item $newFile -Force }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $wb = $excel.Workbooks.Open($targetFile)
    $sh = $wb.Sheets.Item(1)
    
    $startRow = 14
    $c = 3 # Col C (회계재고가)
    
    $r = $startRow
    while ($true) {
        $name = $sh.Cells.Item($r, 2).Value2
        if (-not $name) { break }
        
        $valKrw100M = $sh.Cells.Item($r, $c).Value2
        
        if ($valKrw100M -is [System.ValueType] -or $valKrw100M -match "^[\d\.]+$") {
            $valKrw = $valKrw100M * 100000000
            $valUsd = $valKrw / 1427
            $valUsdM = $valUsd / 1000000
            
            $sh.Cells.Item($r, $c).Value2 = $valUsdM
        }
        $r++
    }
    
    # 다른 이름으로 저장
    $wb.SaveAs($newFile)
    Write-Host "File Saved: $newFile"
    
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
} catch {
    Write-Error $_
}

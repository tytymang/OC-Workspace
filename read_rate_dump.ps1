
$ErrorActionPreference = "Stop"

try {
    # 1. 환율 파일 찾기 (기말_FY2026_환율표_202602.xlsx 추정)
    # 파일명: *환율*.xlsx (Unicode: ȯǥ)
    
    # 20260306_ ǥ ۾ 폴더 내 파일 확인
    $folder = "C:\Users\307984\.openclaw\workspace\Working\20260306_경영 지표 작업" # 실제 폴더명 추정
    
    # Use brute force search
    $files = Get-ChildItem "C:\Users\307984\.openclaw\workspace\Working" -Recurse
    $rateFile = ($files | Where-Object { $_.Name -like "*202602.xlsx" -and $_.Length -gt 200000 }).FullName
    
    if (-not $rateFile) {
        Write-Host "Trying manual match..."
        # From previous ls: C:\Users\307984\.openclaw\workspace\Working\20260306_ ǥ ۾\繫_FY2026_ȯǥ_202602.xlsx
        $rateFile = (Get-ChildItem "C:\Users\307984\.openclaw\workspace\Working" -Recurse | Where-Object { $_.Name -match "202602.xlsx" }).FullName
    }
    
    Write-Host "Rate File: $rateFile"
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $wb = $excel.Workbooks.Open($rateFile)
    $sh = $wb.Sheets.Item(1)
    
    # Dump 1-20 rows, 1-20 cols to find rates
    for ($r=1; $r -le 20; $r++) {
        $rowStr = ""
        for ($c=1; $c -le 20; $c++) {
            $val = $sh.Cells.Item($r, $c).Value2
            if ($val) { $rowStr += "[$r,$c]: $val | " }
        }
        if ($rowStr) { Write-Host $rowStr }
    }
    
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
} catch {
    Write-Error $_
}

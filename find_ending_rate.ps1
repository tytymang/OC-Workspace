
$ErrorActionPreference = "Stop"

try {
    # 1. 파일 경로 직접 지정 (Unicode 문제 방지)
    $files = Get-ChildItem -Path "C:\Users\307984\.openclaw\workspace\Working" -Recurse -Include *202602*.xlsx
    $rateFile = ($files | Where-Object { $_.Name -like "*ȯ*" -or $_.Name -like "*exchange*" -or $_.Name -like "*rate*" })[0].FullName
    
    # 만약 위의 필터로 못 찾으면 파일 크기나 다른 조건으로
    if (-not $rateFile) {
         # 가장 마지막 파일 선택 (기말_FY2026_환율표_202602.xlsx)
         $rateFile = $files[-1].FullName
    }
    
    Write-Host "Target Rate File: $rateFile"
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $wb = $excel.Workbooks.Open($rateFile)
    
    # 2. 첫 번째 시트의 기말 환율 찾기
    $sh = $wb.Sheets.Item(1)
    
    # 기말환율은 보통 특정 셀에 있음.
    # 이전 덤프에서 Row 2에 "[기말환율]" 헤더가 있었음.
    # 그리고 월별 데이터는 컬럼으로 나열될 가능성 높음.
    
    # 헤더 찾기
    $foundEnding = $false
    $endingRow = 0
    $endingCol = 0 # "기말환율"이 시작되는 컬럼
    
    # 1~10행 스캔
    for ($r=1; $r -le 10; $r++) {
        for ($c=1; $c -le 20; $c++) {
             $val = $sh.Cells.Item($r, $c).Value2
             if ($val -ne $null -and ($val -match "기말" -or $val -match "Ending")) {
                 Write-Host "Found Ending Header at [$r, $c]: $val"
                 $endingRow = $r
                 $endingCol = $c
                 $foundEnding = $true
                 break
             }
        }
        if ($foundEnding) { break }
    }
    
    # 월 찾기 (2026.01, 2026.02 등)
    # 헤더 행 또는 그 아래 행에 월 정보가 있을 것임.
    
    if ($foundEnding) {
        $searchRow = $endingRow + 1 # 보통 헤더 바로 아래나 같은 줄 옆
        for ($c=1; $c -le 20; $c++) {
            $val = $sh.Cells.Item($searchRow, $c).Value2
            if ($val) { Write-Host "Row $searchRow Col $c: $val" }
            
            # 혹시 날짜 형식일 수 있음
             $val2 = $sh.Cells.Item($searchRow + 1, $c).Value2
             if ($val2) { Write-Host "Row $($searchRow+1) Col $c: $val2" }
        }
    } else {
        Write-Host "Header '기말' not found via exact match. Dumping rows 1-5 again."
        for ($r=1; $r -le 5; $r++) {
            $rowStr = ""
            for ($c=1; $c -le 20; $c++) {
                $v = $sh.Cells.Item($r, $c).Value2
                if ($v) { $rowStr += "[$r,$c]: $v | " }
            }
            Write-Host $rowStr
        }
    }
    
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
} catch {
    Write-Error $_
}

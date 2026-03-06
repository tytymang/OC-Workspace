
$ErrorActionPreference = "Stop"

try {
    # 1. 파일 경로 설정
    $targetFile = "C:\Users\307984\.openclaw\workspace\Working\20260306_경영 지표 작업\20260205_01월 경영 지표 Report_3.xlsx"
    $newFile = "C:\Users\307984\.openclaw\workspace\Working\20260306_01월말 재고 꼬리표 Rpeort_USD.xlsx"
    
    # PowerShell에서 경로 한글 깨짐 방지용 와일드카드 사용
    $files = Get-ChildItem -Path "C:\Users\307984\.openclaw\workspace\Working" -Recurse -Filter "20260205_01*.xlsx"
    if ($files.Count -eq 0) { throw "File not found" }
    $realPath = $files[0].FullName
    
    # 2. 엑셀 열기
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false # 경고창 무시
    
    $workbook = $excel.Workbooks.Open($realPath)
    $sheet = $workbook.Sheets.Item(1)
    
    # 3. 데이터 변환 (억원 -> M$, 환율 1427원)
    # 헤더는 R13, C3 (회계재고가)
    # 데이터는 R14부터 시작
    
    $startRow = 14
    $col = 3 # C열 (회계재고가)
    
    # 데이터가 끝날 때까지 반복 (이름이 없는 행까지)
    $r = $startRow
    while ($true) {
        # 이름 컬럼(B열, 2번째) 확인
        $name = $sheet.Cells.Item($r, 2).Value2
        if ([string]::IsNullOrWhiteSpace($name)) {
            break # 이름 없으면 종료
        }
        
        # 회계재고가 값 읽기 (억원 단위 가정)
        $valKrw100M = $sheet.Cells.Item($r, $col).Value2
        
        if ($valKrw100M -is [System.ValueType]) { # 숫자인 경우만
            # 억원 -> 원
            $valKrw = $valKrw100M * 100000000
            
            # 원 -> 달러 (환율 1427)
            $valUsd = $valKrw / 1427
            
            # 달러 -> M$ (백만달러)
            $valUsdM = $valUsd / 1000000
            
            # 업데이트
            $sheet.Cells.Item($r, $col).Value2 = $valUsdM
            
            Write-Host "Row $r: $valKrw100M (100M KRW) -> $valUsdM (M$)"
        }
        
        $r++
    }
    
    # 4. 다른 이름으로 저장
    $workbook.SaveAs($newFile)
    Write-Host "Saved as: $newFile"
    
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
} catch {
    Write-Error $_
}

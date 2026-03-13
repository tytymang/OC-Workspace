
$ErrorActionPreference = "Stop"

try {
    # 1. 파일 찾기 (01월 경영 지표)
    # 20260205_01월 경영 지표 Report_3.xlsx 가 가장 유력해 보임
    $files = Get-ChildItem -Path "C:\Users\307984\.openclaw\workspace\Working" -Recurse -Filter "*01*Report*.xlsx"
    if ($files.Count -eq 0) {
        throw "Files not found."
    }
    $targetFile = $files[0].FullName
    Write-Host "Target File: $targetFile"
    
    # 새 파일 경로
    $newFile = "C:\Users\307984\.openclaw\workspace\Working\20260306_01월말 재고 꼬리표 Rpeort_USD.xlsx"

    # 2. 엑셀 열기
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($targetFile)
    $sheet = $workbook.Sheets.Item(1)
    Write-Host "Sheet Name: $($sheet.Name)"

    # 3. 데이터 검색 및 전체 수정
    # 1월말 환율 적용 (1427원)
    $rate = 1427
    
    # "회계재고가" 헤더 찾기 (유니코드)
    $searchHeader = [char]0xD68C + [char]0xACC4 + [char]0xC7AC + [char]0xACE0 + [char]0xAC00 # 회계재고가
    $searchHeader2 = "재고가" # 혹시 모르니
    
    $headerRow = 0
    $headerCol = 0
    
    # 헤더 찾기 (1~20행)
    for ($r=1; $r -le 20; $r++) {
        for ($c=1; $c -le 20; $c++) {
            $val = $sheet.Cells.Item($r, $c).Value2
            if ($val -ne $null -and $val.ToString().Contains($searchHeader)) {
                $headerRow = $r
                $headerCol = $c
                Write-Host "Header found at R$($r) C$($c)"
                break
            }
        }
        if ($headerRow -gt 0) { break }
    }
    
    if ($headerRow -eq 0) { throw "Header not found" }
    
    # 데이터 행 반복 수정
    # 헤더 아래 행부터 값이 있는 동안 계속
    $currentRow = $headerRow + 1
    
    while ($true) {
        # 이름 컬럼(보통 헤더 왼쪽이나 같은 줄의 첫번째)에 값이 없으면 종료로 간주
        # 여기서는 회계재고가 컬럼의 값을 확인
        
        # 이름 확인 (컬럼 2번 가정, 이전 덤프 참조)
        $nameVal = $sheet.Cells.Item($currentRow, 2).Value2
        if ($nameVal -eq $null -or $nameVal -eq "") {
            # 혹시 중간에 빈 줄이 있을 수 있으니 3줄 더 확인
            $emptyCount = 0
            for ($i=1; $i -le 3; $i++) {
                if ($sheet.Cells.Item($currentRow + $i, 2).Value2 -eq $null) {
                    $emptyCount++
                }
            }
            if ($emptyCount -eq 3) { break }
        }
        
        # 값 읽기
        $val = $sheet.Cells.Item($currentRow, $headerCol).Value2
        
        # 값이 숫자이고 0보다 크면 변환
        # 기존 값이 억원 단위라고 가정 (240.409 처럼)
        # 하지만 M$로 이미 되어 있을 수도 있음.
        # 파일 내용을 먼저 확인해야 함. 일단 1월 파일도 덤프해서 확인 필요.
        
        if ($val -is [System.ValueType] -and $val -gt 0) {
            # 1월 파일의 단위 확인이 중요함.
            # 일단 억원 단위로 가정하고 변환 로직 적용
            # 240.409 (억원) -> 24040900000 (원) -> / 1427 (달러) -> / 1000000 (M$)
            
            # 주의: 이미 M$인 경우 엉망이 됨.
            # 1월 파일의 헤더나 값을 먼저 확인하는게 안전함.
            
            # 일단 여기서는 스크립트를 분리해서 확인 먼저 수행.
            break 
        }
        
        $currentRow++
    }
    
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
} catch {
    Write-Error $_
}

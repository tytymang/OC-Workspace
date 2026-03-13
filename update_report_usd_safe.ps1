$ErrorActionPreference = "Stop"

try {
    # 1. 파일 찾기
    $files = Get-ChildItem -Path "C:\Users\307984\.openclaw\workspace\Working" -Recurse -Filter "*20260305_02*.xlsx"
    if ($files.Count -eq 0) {
        Write-Error "Files not found."
    }
    $targetFile = $files[0].FullName
    Write-Host "Target File: $targetFile"

    # 2. 엑셀 열기
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($targetFile)
    $sheet = $workbook.Sheets.Item(1) # First sheet ("보고서")
    Write-Host "Sheet Name: $($sheet.Name)"

    # 3. 데이터 검색
    # 한글 깨짐 방지를 위해 유니코드 코드포인트 사용하거나, 
    # Find 메소드에 직접 문자열을 넣기보다 다른 방법을 써야 할 수도 있음.
    # 하지만 PowerShell ISE에서는 잘 되는데 여기서는 문제임.
    # 일단 영어로 메시지를 출력하고, 검색어는 유니코드 이스케이프로 시도하거나 
    # 그냥 '선봉규'를 검색해봄. (깨질 가능성 높음)
    
    # 선봉규 -> [char]0xC120 + [char]0xBD09 + [char]0xADDC (대충 이런식)
    # 하지만 복잡하니, 셀 값을 하나씩 읽어서 비교하는 방식으로 변경.
    
    $foundName = $null
    $foundHeader = $null
    
    # 헤더 찾기 (1~5행 검색)
    for ($r=1; $r -le 5; $r++) {
        for ($c=1; $c -le 20; $c++) {
            $val = $sheet.Cells.Item($r, $c).Value2
            if ($val -like "*회계재고가*") {
                $foundHeader = @{ Row = $r; Column = $c }
                break
            }
        }
        if ($foundHeader) { break }
    }
    
    if (-not $foundHeader) { Write-Error "Header not found" }
    Write-Host "Header found at R$($foundHeader.Row) C$($foundHeader.Column)"

    # 이름 찾기 (헤더 아래 20행 검색)
    for ($r=$foundHeader.Row + 1; $r -le $foundHeader.Row + 20; $r++) {
        # 이름 컬럼이 보통 A열이나 B열임. 1~5열 검색
        for ($c=1; $c -le 5; $c++) {
            $val = $sheet.Cells.Item($r, $c).Value2
            if ($val -eq "선봉규") {
                $foundName = @{ Row = $r; Column = $c }
                break
            }
        }
        if ($foundName) { break }
    }
    
    if (-not $foundName) { Write-Error "Name not found" }
    Write-Host "Name found at R$($foundName.Row) C$($foundName.Column)"

    # 4. 값 읽기 및 계산
    $targetCell = $sheet.Cells.Item($foundName.Row, $foundHeader.Column)
    $currentValue = $targetCell.Value2
    Write-Host "Current Value (100M KRW): $currentValue"

    # 환율 적용 (1427원)
    $exchangeRate = 1427
    $krwValue = $currentValue * 100000000
    $usdValue = $krwValue / $exchangeRate
    
    Write-Host "Calculated USD: $usdValue"

    # 5. 값 업데이트
    $targetCell.Value2 = $usdValue
    
    # 저장
    $workbook.Save()
    Write-Host "File Saved."

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { 
        $excel.Quit() 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

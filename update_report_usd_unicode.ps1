
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

    # 3. 데이터 검색 (유니코드 사용)
    $searchName = [char]0xC120 + [char]0xBD09 + [char]0xADDC # 선봉규
    $searchHeader = [char]0xD68C + [char]0xACC4 + [char]0xC7AC + [char]0xACE0 + [char]0xAC00 # 회계재고가
    
    $foundName = $null
    $foundHeader = $null
    
    $usedRange = $sheet.UsedRange
    $rows = $usedRange.Rows.Count
    $cols = $usedRange.Columns.Count

    # 헤더 찾기 (1~10행 검색)
    for ($r=1; $r -le 10; $r++) {
        for ($c=1; $c -le 20; $c++) {
            $val = $sheet.Cells.Item($r, $c).Value2
            if ($val -match $searchHeader) {
                $foundHeader = @{ Row = $r; Column = $c }
                Write-Host "Header found at R$($foundHeader.Row) C$($foundHeader.Column)"
                break
            }
        }
        if ($foundHeader) { break }
    }
    
    if (-not $foundHeader) { Write-Error "Header ($searchHeader) not found" }
    
    # 이름 찾기 (헤더 행부터 끝까지 검색)
    for ($r=$foundHeader.Row; $r -le $rows; $r++) {
        # 이름 컬럼이 보통 A~E열임.
        for ($c=1; $c -le 5; $c++) {
            $val = $sheet.Cells.Item($r, $c).Value2
            if ($val -eq $searchName) {
                $foundName = @{ Row = $r; Column = $c }
                Write-Host "Name found at R$($foundName.Row) C$($foundName.Column)"
                break
            }
        }
        if ($foundName) { break }
    }
    
    if (-not $foundName) { Write-Error "Name ($searchName) not found" }

    # 4. 값 읽기 및 계산
    # 회계재고가 열에서 해당 행의 값 읽기
    $targetRow = $foundName.Row
    $targetCol = $foundHeader.Column
    
    $targetCell = $sheet.Cells.Item($targetRow, $targetCol)
    $currentValue = $targetCell.Value2
    Write-Host "Current Value (100M KRW): $currentValue"

    if ($currentValue -eq $null -or $currentValue -eq 0) {
        Write-Error "Value is null or zero."
    }

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

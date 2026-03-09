
$ErrorActionPreference = "Stop"

try {
    # 1. Find File
    $files = Get-ChildItem -Path "C:\Users\307984\.openclaw\workspace\Working" -Recurse -Filter "*20260305_02*.xlsx"
    $targetFile = $files[0].FullName
    
    # 2. Open Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($targetFile)
    $sheet = $workbook.Sheets.Item(1)
    
    # 3. Find Name (선봉규)
    $searchName = [char]0xC120 + [char]0xBD09 + [char]0xADDC
    $found = $sheet.UsedRange.Find($searchName)
    
    if ($found) {
        $r = $found.Row
        $c = 3 # 회계재고가(M$) 컬럼
        
        # 현재 값 읽기
        $val = $sheet.Cells.Item($r, $c).Value2
        Write-Host "Current Value: $val (M$)"
        
        # 240.409 억원을 USD로 변환
        # 240.409 억원 = 24,040,900,000 원
        # 1월 환율 1427원 적용
        $krw = 24040900000
        $rate = 1427
        $usd = $krw / $rate # 달러
        $usd_m = $usd / 1000000 # M$
        
        Write-Host "New Value (from 240.409 100M KRW): $usd_m (M$)"
        
        # 값 업데이트
        $sheet.Cells.Item($r, $c).Value2 = $usd_m
        
        # 저장
        $workbook.Save()
        Write-Host "Saved."
    } else {
        Write-Host "Name not found"
    }

} catch {
    Write-Error $_
} finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { 
        $excel.Quit() 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

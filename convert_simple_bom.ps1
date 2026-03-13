# BOM 추가를 위해 한글 주석 포함
# 이 파일은 UTF-8 with BOM 으로 저장되어야 함

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$path = "C:\Users\307984\.openclaw\workspace\temp_test\20260205_01월말 재고 꼬리표 Report_3.xlsx"

try {
    $wb = $excel.Workbooks.Open($path)
    $found = $false
    
    foreach ($sh in $wb.Sheets) {
        # 전체 범위 검색
        $ur = $sh.UsedRange
        $range = $sh.Range("A1", $sh.Cells.Item($ur.Rows.Count, $ur.Columns.Count))
        
        # 선봉규 찾기
        $cell = $range.Find("선봉규")
        
        if ($cell) {
            $r = $cell.Row
            
            # P열(16)이 KRW, N열(14)이 USD라고 가정 (헤더 검색 생략)
            $krw = $sh.Cells.Item($r, 16).Value2
            
            if ($krw -is [double] -or $krw -is [int]) {
                # 억원 -> 원 -> 달러 (환율 1427)
                $usd = ($krw * 100000000) / 1427
                
                # N열에 입력
                $sh.Cells.Item($r, 14).Value2 = $usd
                $found = $true
            }
            break
        }
    }
    
    if ($found) {
        $wb.Save()
        Write-Output "SUCCESS"
    }

} catch {
    # Error
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

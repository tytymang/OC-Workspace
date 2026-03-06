
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
        # Find 메서드 대신 행 단위 스캔 (한글 문자열 '선봉규' 매칭 문제 회피)
        
        $rowCount = $ur.Rows.Count
        $colCount = $ur.Columns.Count
        
        # 데이터가 있는 영역 (예: 6행부터) 스캔
        for ($r = 6; $r -le $rowCount; $r++) {
            # B열(2) ~ F열(6) 사이에서 이름 찾기 (보통 이름은 앞쪽에 있음)
            for ($c = 2; $c -le 6; $c++) {
                $val = $sh.Cells.Item($r, $c).Text
                if ($val -match "선봉규") {
                    
                    # P열(16) KRW 읽기
                    $krw = $sh.Cells.Item($r, 16).Value2
                    
                    if ($krw -is [double] -or $krw -is [int]) {
                        # 계산 (환율 1427)
                        $usd = ($krw * 100000000) / 1427
                        
                        # N열(14)에 쓰기
                        $sh.Cells.Item($r, 14).Value2 = $usd
                        $found = $true
                        break
                    }
                }
            }
            if ($found) { break }
        }
        if ($found) { break }
    }
    
    if ($found) {
        $wb.Save()
        Write-Output "SUCCESS"
    }

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

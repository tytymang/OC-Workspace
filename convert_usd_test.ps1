# Korean Encoding Skill Applied
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 경로 설정 (temp_test)
$basePath = "C:\Users\307984\.openclaw\workspace\temp_test"
$targetFile = Get-ChildItem -Path $basePath -Recurse -Filter "*01월말*Report*.xlsx" | Select-Object -First 1

if (!$targetFile) { Write-Error "File not found"; exit }
$path = $targetFile.FullName
Write-Host "Editing: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    $found = $false
    $exchRate = 1427.00
    
    # 모든 시트 검색
    foreach ($sheet in $workbook.Sheets) {
        $usedRange = $sheet.UsedRange
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count
        
        # '선봉규' 찾기 (Find 메서드 사용)
        $searchRange = $sheet.Range("A1", $sheet.Cells.Item($rowCount, $colCount))
        $foundCell = $searchRange.Find("선봉규")
        
        if ($foundCell) {
            $row = $foundCell.Row
            Write-Host "Found '선봉규' in Sheet '$($sheet.Name)' at Row $row"
            
            # 컬럼 매핑 (헤더 검색)
            $krwCol = 0
            $usdCol = 0
            
            # 헤더 행 추정 (데이터 위쪽 1~5행 스캔)
            for ($h = 1; $h -le $row; $h++) {
                for ($c = 1; $c -le $colCount; $c++) {
                    $val = $sheet.Cells.Item($h, $c).Text
                    if ($val -match "회계" -and ($val -match "재고" -or $val -match "금액")) { $krwCol = $c }
                    if ($val -match "USD" -or $val -match "외화") { $usdCol = $c }
                }
                # 헤더 찾으면 중단 (가장 가까운 헤더 사용)
                if ($krwCol -gt 0) { break }
            }
            
            # 못 찾았으면 수동 매핑 (보통 P열이 KRW, N열이 USD였음)
            if ($krwCol -eq 0) { $krwCol = 16 } # P열 가정
            if ($usdCol -eq 0) { $usdCol = 14 } # N열 가정 (없으면 P열 왼쪽에 기록)
            
            # 값 읽기 (억원 단위)
            $krwVal = $sheet.Cells.Item($row, $krwCol).Value2
            
            if ($krwVal -is [double] -or $krwVal -is [int]) {
                # 계산: (억원 * 1억) / 환율
                $realKRW = $krwVal * 100000000
                $usdVal = $realKRW / $exchRate
                
                Write-Host "Original KRW (100M Unit): $krwVal"
                Write-Host "Real KRW: $realKRW"
                Write-Host "Calculated USD: $usdVal (Rate: $exchRate)"
                
                # USD 컬럼에 기록 (없으면 KRW 옆에 비고로?)
                # 지시: "USD로 변환해봐" -> USD 컬럼에 넣거나 KRW를 덮어쓰기?
                # "오직 ... 회계 재고가만 수정해봐" -> USD 컬럼에 넣는게 안전.
                
                if ($usdCol -gt 0) {
                    $sheet.Cells.Item($row, $usdCol).Value2 = $usdVal
                    Write-Host "Updated Cell [${row}, ${usdCol}] with USD value."
                } else {
                    Write-Warning "USD Column not found. Overwriting KRW cell (Risk!) or adding comment?"
                    # 안전하게 바로 옆 셀(오른쪽)에 씀
                    $sheet.Cells.Item($row, $krwCol + 1).Value2 = $usdVal
                    Write-Host "Wrote to Right Cell [${row}, $($krwCol+1)]"
                }
                $found = $true
            } else {
                Write-Warning "KRW value is not a number: $krwVal"
            }
            break # 한 명만 수정
        }
    }
    
    if ($found) {
        $workbook.Save()
        Write-Host "SUCCESS: Saved changes."
    } else {
        Write-Warning "Target '선봉규' not found."
    }

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

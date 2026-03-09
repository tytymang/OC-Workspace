# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 1. 대상 파일 정의 (temp_report 폴더)
$baseDir = "C:\Users\307984\.openclaw\workspace\temp_report"

# 파일 찾기
$targetFile = Get-ChildItem -Path $baseDir -Recurse -Filter "*20260305_02*.xlsx" | Select-Object -First 1

if (!$targetFile) {
    Write-Error "Target file not found."
    exit
}

$path = $targetFile.FullName
Write-Host "Applying Rate 1449.32 to: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    $sheet = $workbook.Sheets.Item(3) # 3번째 시트 (Main Data)

    # 2. 환율 업데이트 (P2 셀)
    # 기존에 1,290.93이 있던 자리
    $targetRate = 1449.32
    $sheet.Cells.Item(2, 16).Value2 = $targetRate
    Write-Host "Updated P2 Rate: $targetRate"

    # 3. 데이터 행에 수식 적용
    $usedRange = $sheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    
    # 데이터 시작 행: 6행 (아까 확인)
    $startRow = 6
    if ($rowCount -ge $startRow) {
        # P열 (16열) = N열 (14열) * $P$2
        # Q열 (17열) = O열 (15열) * $P$2
        
        # 범위 지정 (P6:Pxxxx)
        $rangeP = $sheet.Range("P${startRow}:P${rowCount}")
        $rangeP.Formula = "=N${startRow}*`$P`$2" # P2 절대참조
        
        # 범위 지정 (Q6:Qxxxx)
        $rangeQ = $sheet.Range("Q${startRow}:Q${rowCount}")
        $rangeQ.Formula = "=O${startRow}*`$P`$2" 
        
        Write-Host "Applied formula to Rows ${startRow}~${rowCount}"
    }

    $workbook.Save()
    Write-Host "SUCCESS: Formula applied."

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

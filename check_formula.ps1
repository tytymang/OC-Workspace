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
Write-Host "Checking Formula in P6: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    $sheet = $workbook.Sheets.Item(3) # 3번째 시트

    # P6 (16열, 6행) 수식 확인
    $formula = $sheet.Cells.Item(6, 16).Formula
    Write-Host "Formula in P6: $formula"
    
    # Q6 (17열, 6행) 수식 확인 (혹시 여기도?)
    $formulaQ = $sheet.Cells.Item(6, 17).Formula
    Write-Host "Formula in Q6: $formulaQ"

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 1. 대상 파일 정의 (temp_check 폴더)
$baseDir = "C:\Users\307984\.openclaw\workspace\temp_check"

# 파일 찾기
$targetFile = Get-ChildItem -Path $baseDir -Recurse -Filter "*FY2026*.xlsx" | Select-Object -First 1

if (!$targetFile) {
    Write-Error "Target file not found."
    exit
}

$path = $targetFile.FullName

try {
    $workbook = $excel.Workbooks.Open($path)
    $sheet = $workbook.Sheets.Item("2025")
    
    # 기말환율(3열) 데이터 추출
    # 2026년 1월 (Row 6)
    $jan26 = $sheet.Cells.Item(6, 3).Text
    
    # 2026년 2월 (Row 7)
    $feb26 = $sheet.Cells.Item(7, 3).Text
    
    # 2025년 12월 (Row 35 근처)
    # 아까 35행에서 12월 데이터를 봤으므로, 35행 3열 확인
    $dec25 = $sheet.Cells.Item(35, 3).Text
    
    # 혹시 몰라 34~36행 확인
    if ($sheet.Cells.Item(34, 2).Text -match "12") { $dec25 = $sheet.Cells.Item(34, 3).Text }
    if ($sheet.Cells.Item(36, 2).Text -match "12") { $dec25 = $sheet.Cells.Item(36, 3).Text }

    Write-Host "DEC_2025: $dec25"
    Write-Host "JAN_2026: $jan26"
    Write-Host "FEB_2026: $feb26"

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

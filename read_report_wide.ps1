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
Write-Host "Scanning Header & Right Columns: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    $sheet = $workbook.Sheets.Item(1)

    # 1행 ~ 5행, 1열 ~ 20열 덤프 (오른쪽까지 확인)
    for ($r = 1; $r -le 5; $r++) {
        $rowText = ""
        for ($c = 1; $c -le 20; $c++) { 
            $val = $sheet.Cells.Item($r, $c).Text
            if ($val) { # 값이 있는 경우만
                $rowText += "[${c}:$val] "
            }
        }
        Write-Host "Row ${r}: $rowText"
    }
    
    # 시트 목록도 확인
    Write-Host "--- Sheets ---"
    foreach ($s in $workbook.Sheets) {
        Write-Host "Sheet: $($s.Name)"
    }

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# 엑셀 COM 객체 생성
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 절대 경로 사용
$baseDir = "C:\Users\307984\.openclaw\workspace\temp_work"

# 파일 찾기
$targetFile = Get-ChildItem -Path $baseDir -Recurse -Filter "*FY2026*.xlsx" | Select-Object -First 1

if (!$targetFile) {
    # 서브디렉토리 순회
    $subDirs = Get-ChildItem -Path $baseDir -Directory
    foreach ($d in $subDirs) {
        $f = Get-ChildItem -Path $d.FullName -Filter "*FY2026*.xlsx" | Select-Object -First 1
        if ($f) {
            $targetFile = $f
            break
        }
    }
}

if (!$targetFile) {
    Write-Error "File not found."
    exit
}

$path = $targetFile.FullName
Write-Host "Reading last rows of: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    $sheet = $workbook.Sheets.Item(1) # 첫 번째 시트 (월별분기별)

    $usedRange = $sheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    
    Write-Host "Total Rows: $rowCount"
    
    # 마지막 30행 스캔
    $startRow = $rowCount - 30
    if ($startRow -lt 1) { $startRow = 1 }
    
    for ($r = $startRow; $r -le $rowCount; $r++) {
        $rowText = ""
        for ($c = 1; $c -le 10; $c++) { # 10열까지
            $val = $sheet.Cells.Item($r, $c).Text
            $rowText += "[$val] "
        }
        
        # ${r} 문법 사용
        Write-Host "Row ${r}: $rowText"
    }

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    # 임시 폴더 삭제
    Remove-Item $baseDir -Recurse -Force
}

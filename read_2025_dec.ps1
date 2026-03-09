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
Write-Host "Scanning 2025 Sheet for Dec 2025: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    $sheet = $workbook.Sheets.Item("2025")
    
    # 20행 ~ 50행 덤프 (2025년 12월 찾기)
    for ($r = 20; $r -le 50; $r++) {
        $rowText = ""
        # 날짜 후보(2열)와 USD(26열)
        
        $dateVal = $sheet.Cells.Item($r, 2).Text
        $usdVal = $sheet.Cells.Item($r, 26).Text
        
        if ($dateVal -match "12" -or $dateVal -match "2025") {
            Write-Host "Row ${r}: Date=[$dateVal] USD=[$usdVal]"
        }
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

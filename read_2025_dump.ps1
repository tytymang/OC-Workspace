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
Write-Host "Reading 2025 Sheet content: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    
    # 2025 시트 찾기
    $sheet = $workbook.Sheets.Item("2025")
    
    # 1행 ~ 20행, 1열 ~ 30열 내용 덤프
    # 특히 26열(USD) 주변과 1~5열(날짜 후보) 집중 확인
    for ($r = 1; $r -le 20; $r++) {
        $rowText = ""
        # 1~5열 (날짜 후보)
        for ($c = 1; $c -le 5; $c++) {
            $val = $sheet.Cells.Item($r, $c).Text
            $rowText += "[${c}:$val] "
        }
        
        # 26열 (USD)
        $usdVal = $sheet.Cells.Item($r, 26).Text
        $rowText += "... [26:$usdVal]"
        
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

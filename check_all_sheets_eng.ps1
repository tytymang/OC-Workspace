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
Write-Host "Checking sheets in: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    
    # 모든 시트 순회
    foreach ($sheet in $workbook.Sheets) {
        Write-Host "Sheet: $($sheet.Name)"
        
        # 1~30행 덤프 (USD, 2025, 2026 키워드 찾기)
        for ($r = 1; $r -le 30; $r++) { 
            $rowText = ""
            $foundKeyword = $false
            
            for ($c = 1; $c -le 10; $c++) { # 10열까지 검사
                $val = $sheet.Cells.Item($r, $c).Text
                $rowText += "[$val] "
                
                # 영어/숫자만 사용
                if ($val -match "2025" -or $val -match "2026" -or $val -match "USD") {
                    $foundKeyword = $true
                }
            }
            
            if ($foundKeyword) {
                # ${r} 문법 사용
                Write-Host "  Row ${r}: $rowText"
            }
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

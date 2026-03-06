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
Write-Host "Checking for '2005' Sheet in: $path"

try {
    $workbook = $excel.Workbooks.Open($path)
    
    # 2005 시트 찾기
    $sheet = $null
    foreach ($s in $workbook.Sheets) {
        if ($s.Name -eq "2005") {
            $sheet = $s
            break
        }
    }
    
    if ($sheet) {
        Write-Host "FOUND Sheet: 2005"
        # 1행 ~ 20행, 1열 ~ 10열 덤프
        for ($r = 1; $r -le 20; $r++) {
            $rowText = ""
            for ($c = 1; $c -le 10; $c++) { 
                $val = $sheet.Cells.Item($r, $c).Text
                $rowText += "[${c}:$val] "
            }
            Write-Host "Row ${r}: $rowText"
        }
    } else {
        Write-Warning "Sheet '2005' NOT found."
        Write-Host "Available Sheets:"
        foreach ($s in $workbook.Sheets) {
            Write-Host "- $($s.Name)"
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

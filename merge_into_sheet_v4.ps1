
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$workingDir = "C:\Users\307984\.openclaw\workspace\working"
$outputFile = Join-Path $workingDir "AI KPI_20260226.xlsx"

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $allFiles = Get-ChildItem -Path $workingDir -Filter "*.xlsx"
    $templateFileObj = $allFiles | Where-Object { $_.Name -like "20260126_*" }
    
    $targetWb = $excel.Workbooks.Open($templateFileObj.FullName)
    $targetSheet = $null
    foreach ($sh in $targetWb.Sheets) { if ($sh.Name -like "*AI KPI*") { $targetSheet = $sh; break } }
    
    $sourceFiles = $allFiles | Where-Object { 
        $_.Name -ne $templateFileObj.Name -and $_.Name -ne "AI KPI_20260226.xlsx" 
    }

    foreach ($file in $sourceFiles) {
        Write-Output "Processing: $($file.Name)"
        $sourceWb = $excel.Workbooks.Open($file.FullName)
        $sourceSheet = $sourceWb.Sheets.Item(1)
        
        $sourceLastRow = $sourceSheet.UsedRange.Rows.Count
        if ($sourceLastRow -ge 2) {
            $sourceData = $sourceSheet.Range("A2", $sourceSheet.Cells.Item($sourceLastRow, 20)).Value2
            
            # 타겟 시트 마지막 행 계산
            $targetLastRow = $targetSheet.UsedRange.Rows.Count
            # 만약 UsedRange가 비어있어도 1로 나오므로 실제 데이터 확인 필요하지만, 보통 제목행이 있음
            $destRange = $targetSheet.Range($targetSheet.Cells.Item($targetLastRow + 1, 1), $targetSheet.Cells.Item($targetLastRow + ($sourceLastRow - 1), 20))
            $destRange.Value2 = $sourceData
        }
        $sourceWb.Close($false)
    }

    if (Test-Path $outputFile) { Remove-Item $outputFile }
    $targetWb.SaveAs($outputFile)
    $targetWb.Close($false)
    Write-Output "SUCCESS"
} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($excel) { $excel.Quit() }
}


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
    
    if (-not $templateFileObj) { throw "Template not found." }
    
    $targetWb = $excel.Workbooks.Open($templateFileObj.FullName)
    $targetSheet = $null
    foreach ($sh in $targetWb.Sheets) {
        if ($sh.Name -like "*AI KPI*") { $targetSheet = $sh; break }
    }
    
    $sourceFiles = $allFiles | Where-Object { 
        $_.Name -ne $templateFileObj.Name -and $_.Name -ne "AI KPI_20260226.xlsx" 
    }

    foreach ($file in $sourceFiles) {
        Write-Output "Copying from: $($file.Name)"
        $sourceWb = $excel.Workbooks.Open($file.FullName)
        $sourceSheet = $sourceWb.Sheets.Item(1)
        
        $sourceLastRow = $sourceSheet.UsedRange.Rows.Count
        if ($sourceLastRow -ge 2) {
            $copyRange = $sourceSheet.Range("A2", $sourceSheet.Cells.Item($sourceLastRow, 20)) # T열까지
            $copyRange.Copy() | Out-Null
            
            # 마지막 행 찾기 (데이터가 있는 실제 마지막 행)
            $targetLastRow = $targetSheet.UsedRange.Rows.Count
            $destCell = $targetSheet.Cells.Item($targetLastRow + 1, 1)
            
            # PasteSpecial 대신 단순 Paste 시도 또는 다른 방식
            $targetSheet.Paste($destCell) | Out-Null
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


[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$workingDir = "C:\Users\307984\.openclaw\workspace\working"
$outputFile = Join-Path $workingDir "AI KPI_20260226.xlsx"

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # 파일명을 패턴으로 찾기
    $allFiles = Get-ChildItem -Path $workingDir -Filter "*.xlsx"
    $templateFileObj = $allFiles | Where-Object { $_.Name -like "20260126_*" }
    
    if (-not $templateFileObj) { throw "Template file starting with '20260126' not found." }
    
    Write-Output "Template found: $($templateFileObj.Name)"
    $targetWb = $excel.Workbooks.Open($templateFileObj.FullName)
    
    # 'AI KPI' 시트 찾기
    $targetSheet = $null
    foreach ($sh in $targetWb.Sheets) {
        if ($sh.Name -like "*AI KPI*") { $targetSheet = $sh; break }
    }
    if (-not $targetSheet) { throw "'AI KPI' sheet not found in template." }

    $sourceFiles = $allFiles | Where-Object { 
        $_.Name -ne $templateFileObj.Name -and $_.Name -ne "AI KPI_20260226.xlsx" 
    }

    foreach ($file in $sourceFiles) {
        Write-Output "Appending content from: $($file.Name)"
        try {
            $sourceWb = $excel.Workbooks.Open($file.FullName)
            $sourceSheet = $sourceWb.Sheets.Item(1)
            
            # 소스 데이터 범위 (A2부터 마지막 데이터가 있는 곳까지)
            $sourceLastRow = $sourceSheet.UsedRange.Rows.Count
            if ($sourceLastRow -lt 2) { 
                $sourceWb.Close($false)
                continue 
            }

            # 타겟 시트의 마지막 행 계산
            $targetLastRow = $targetSheet.UsedRange.Rows.Count
            $destRow = $targetLastRow + 1

            # 복사 (Z열까지 넉넉하게 잡음)
            $copyRange = $sourceSheet.Range("A2", "Z" + $sourceLastRow)
            $copyRange.Copy()
            
            $destRange = $targetSheet.Cells.Item($destRow, 1)
            $destRange.PasteSpecial(-4163) # xlPasteValues
            
            $sourceWb.Close($false)
        } catch {
            Write-Warning "Failed to append $($file.Name): $($_.Exception.Message)"
        }
    }

    if (Test-Path $outputFile) { Remove-Item $outputFile }
    $targetWb.SaveAs($outputFile)
    $targetWb.Close($false)
    Write-Output "SUCCESS: Integrated all contents into '$($targetSheet.Name)' sheet."

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

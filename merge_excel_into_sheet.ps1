
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$workingDir = "C:\Users\307984\.openclaw\workspace\working"
$templateFile = "20260126_부문별 KPI 집계_작성중.xlsx"
$outputFile = Join-Path $workingDir "AI KPI_20260226.xlsx"

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # 1. 템플릿 파일 열기
    $templatePath = Join-Path $workingDir $templateFile
    if (-not (Test-Path $templatePath)) { throw "Template not found: $templatePath" }
    
    $targetWb = $excel.Workbooks.Open($templatePath)
    $targetSheet = $null
    
    # 'AI KPI' 시트 찾기
    foreach ($sh in $targetWb.Sheets) {
        if ($sh.Name -eq "AI KPI") { $targetSheet = $sh; break }
    }
    if (-not $targetSheet) { throw "'AI KPI' sheet not found in template." }

    # 2. 소스 파일들 (템플릿 및 결과 파일 제외)
    $sourceFiles = Get-ChildItem -Path $workingDir -Filter "*.xlsx" | Where-Object { 
        $_.Name -ne $templateFile -and $_.Name -ne "AI KPI_20260226.xlsx" 
    }

    Write-Output "Template 'AI KPI' sheet loaded."

    foreach ($file in $sourceFiles) {
        Write-Output "Merging content from: $($file.Name)"
        try {
            $sourceWb = $excel.Workbooks.Open($file.FullName)
            $sourceSheet = $sourceWb.Sheets.Item(1)
            
            # 데이터가 있는 마지막 행 찾기 (A열 기준 또는 UsedRange)
            $sourceLastRow = $sourceSheet.UsedRange.Rows.Count
            if ($sourceLastRow -lt 2) { 
                $sourceWb.Close($false)
                continue 
            }

            # 제목행(1행) 제외하고 데이터만 복사 (2행부터 시작한다고 가정)
            # 만약 템플릿의 형식을 유지하며 아래로 계속 붙여넣어야 한다면:
            $targetLastRow = $targetSheet.UsedRange.Rows.Count
            if ($targetLastRow -eq 1 -and [string]::IsNullOrWhiteSpace($targetSheet.Cells.Item(1,1).Value)) {
                $destRow = 1
            } else {
                $destRow = $targetLastRow + 1
            }

            # 복사할 범위 지정 (A2부터 데이터 끝까지)
            $copyRange = $sourceSheet.Range("A2", "Z" + $sourceLastRow) # 여유있게 Z열까지
            $copyRange.Copy()
            
            $destRange = $targetSheet.Cells.Item($destRow, 1)
            $destRange.PasteSpecial(-4163) # xlPasteValues (값만 붙여넣기)
            
            $sourceWb.Close($false)
        } catch {
            Write-Warning "Failed: $($file.Name) - $($_.Exception.Message)"
        }
    }

    # 3. 저장
    if (Test-Path $outputFile) { Remove-Item $outputFile }
    $targetWb.SaveAs($outputFile)
    $targetWb.Close($false)
    Write-Output "SUCCESS: Integrated all contents into 'AI KPI' sheet at $outputFile"

} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

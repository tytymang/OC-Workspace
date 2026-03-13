$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $folder = "C:\Users\307984\.openclaw\workspace\temp_attachments"
    $files = Get-ChildItem $folder -Filter "*.xlsx"
    
    $workbook = $excel.Workbooks.Open($files[0].FullName)
    $sheetNames = foreach ($s in $workbook.Sheets) { $s.Name }
    
    $aiData = @()
    # "AI"가 포함된 시트가 있는지 확인
    $aiSheet = $workbook.Sheets | Where-Object { $_.Name -like "*AI*" }
    if ($null -ne $aiSheet) {
        $rowCount = $aiSheet.UsedRange.Rows.Count
        $colCount = $aiSheet.UsedRange.Columns.Count
        for ($r = 1; $r -le $rowCount; $r++) {
            $row = @()
            for ($c = 1; $c -le $colCount; $c++) {
                $row += $aiSheet.Cells.Item($r, $c).Text
            }
            $aiData += ,$row
        }
    }
    
    $workbook.Close($false)
    $excel.Quit()
    
    @{
        Sheets = $sheetNames
        AIData = $aiData
    } | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open("C:\Users\307984\.openclaw\workspace\temp_attachments\2본부_AI.xlsx")
    $sheet = $workbook.Sheets.Item(1)
    
    $results = @()
    $rowCount = $sheet.UsedRange.Rows.Count
    $colCount = $sheet.UsedRange.Columns.Count

    for ($r = 1; $r -le $rowCount; $r++) {
        $row = @()
        for ($c = 1; $c -le $colCount; $c++) {
            $row += $sheet.Cells.Item($r, $c).Text
        }
        $results += ,$row
    }
    
    $workbook.Close($false)
    $excel.Quit()
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
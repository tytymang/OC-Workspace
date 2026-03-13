$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $folder = "C:\Users\307984\.openclaw\workspace\temp_attachments"
    $files = Get-ChildItem $folder -Filter "*.xlsx"
    if ($files.Count -eq 0) { throw "No Excel file found" }
    
    $workbook = $excel.Workbooks.Open($files[0].FullName)
    $sheet = $workbook.Sheets.Item(1)
    
    $results = @()
    $rowCount = [Math]::Min($sheet.UsedRange.Rows.Count, 100)
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
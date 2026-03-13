
$dir = "C:\Users\307984\.openclaw\workspace\temp_attachments_kim"
$files = Get-ChildItem -Path $dir -Filter "*.xlsx"
if ($files.Count -gt 0) {
    $path = $files[0].FullName
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($path)
    
    foreach ($sheet in $workbook.Sheets) {
        Write-Output "--- SHEET: $($sheet.Name) ---"
        $range = $sheet.UsedRange
        $rowCount = $range.Rows.Count
        $colCount = $range.Columns.Count
        for ($r = 1; $r -le $rowCount; $r++) {
            $rowText = ""
            for ($c = 1; $c -le $colCount; $c++) {
                $cellText = $sheet.Cells.Item($r, $c).Text
                if ($cellText -ne "") {
                    $rowText += $cellText + "|"
                }
            }
            if ($rowText -ne "") {
                Write-Output $rowText
            }
        }
    }

    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
} else {
    Write-Output "No xlsx file found"
}

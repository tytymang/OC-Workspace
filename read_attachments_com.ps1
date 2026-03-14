$excelFiles = Get-ChildItem -Path "C:\Users\307984\.openclaw\workspace\temp_attachments\*.xlsx"
$pptFiles = Get-ChildItem -Path "C:\Users\307984\.openclaw\workspace\temp_attachments\*.pptx"

Write-Output "--- EXCEL FILES ---"
if ($excelFiles) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    foreach ($file in $excelFiles) {
        Write-Output "Opening Excel: $($file.FullName)"
        try {
            $wb = $excel.Workbooks.Open($file.FullName)
            foreach ($sheet in $wb.Sheets) {
                Write-Output "`nSheet: $($sheet.Name)"
                $range = $sheet.UsedRange
                $rows = $range.Rows.Count
                $cols = $range.Columns.Count
                
                # Reading first 20 rows and 10 columns for preview
                $rMax = [math]::Min(20, $rows)
                $cMax = [math]::Min(10, $cols)
                
                for ($r = 1; $r -le $rMax; $r++) {
                    $rowStr = ""
                    for ($c = 1; $c -le $cMax; $c++) {
                        $val = $sheet.Cells.Item($r, $c).Text
                        $rowStr += "$val | "
                    }
                    if ($rowStr.Trim() -ne "|") {
                        Write-Output $rowStr
                    }
                }
            }
            $wb.Close($false)
        } catch {
            Write-Output "Error: $_"
        }
    }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Output "`n--- PPT FILES ---"
if ($pptFiles) {
    $ppt = New-Object -ComObject PowerPoint.Application
    # $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

    foreach ($file in $pptFiles) {
        Write-Output "`nOpening PPT: $($file.FullName)"
        try {
            $presentation = $ppt.Presentations.Open($file.FullName, $false, $false, $false)
            for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
                $slide = $presentation.Slides.Item($i)
                Write-Output "Slide $i"
                foreach ($shape in $slide.Shapes) {
                    if ($shape.HasTextFrame) {
                        $text = $shape.TextFrame.TextRange.Text
                        $text = $text -replace "`r`n", " " -replace "`n", " " -replace "`r", " "
                        if ($text.Trim() -ne "") {
                            Write-Output " - $text"
                        }
                    }
                }
            }
            $presentation.Close()
        } catch {
            Write-Output "Error: $_"
        }
    }
    $ppt.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
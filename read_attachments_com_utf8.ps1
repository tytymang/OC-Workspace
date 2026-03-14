[CmdletBinding()]
Param()

$excelFiles = Get-ChildItem -Path "C:\Users\307984\.openclaw\workspace\temp_attachments\*.xlsx"
$pptFiles = Get-ChildItem -Path "C:\Users\307984\.openclaw\workspace\temp_attachments\*.pptx"
$outFile = "C:\Users\307984\.openclaw\workspace\extracted_data.txt"

# Ensure UTF8 encoding for the file output
$OutputEncoding = [System.Text.Encoding]::UTF8

$outData = @()

$outData += "--- EXCEL FILES ---"
if ($excelFiles) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    foreach ($file in $excelFiles) {
        $outData += "Opening Excel: $($file.FullName)"
        try {
            $wb = $excel.Workbooks.Open($file.FullName)
            foreach ($sheet in $wb.Sheets) {
                $outData += "`nSheet: $($sheet.Name)"
                $range = $sheet.UsedRange
                $rows = $range.Rows.Count
                $cols = $range.Columns.Count
                
                $rMax = [math]::Min(50, $rows)
                $cMax = [math]::Min(15, $cols)
                
                for ($r = 1; $r -le $rMax; $r++) {
                    $rowStr = ""
                    for ($c = 1; $c -le $cMax; $c++) {
                        $val = $sheet.Cells.Item($r, $c).Text
                        $rowStr += "$val | "
                    }
                    if ($rowStr.Trim() -ne "|") {
                        $outData += $rowStr
                    }
                }
            }
            $wb.Close($false)
        } catch {
            $outData += "Error: $_"
        }
    }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

$outData += "`n--- PPT FILES ---"
if ($pptFiles) {
    $ppt = New-Object -ComObject PowerPoint.Application
    foreach ($file in $pptFiles) {
        $outData += "`nOpening PPT: $($file.FullName)"
        try {
            $presentation = $ppt.Presentations.Open($file.FullName, $false, $false, $false)
            for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
                $slide = $presentation.Slides.Item($i)
                $outData += "Slide $i"
                foreach ($shape in $slide.Shapes) {
                    if ($shape.HasTextFrame) {
                        $text = $shape.TextFrame.TextRange.Text
                        $text = $text -replace "`r`n", " " -replace "`n", " " -replace "`r", " "
                        if ($text.Trim() -ne "") {
                            $outData += " - $text"
                        }
                    }
                }
            }
            $presentation.Close()
        } catch {
            $outData += "Error: $_"
        }
    }
    $ppt.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
}

[System.IO.File]::WriteAllLines($outFile, $outData, [System.Text.Encoding]::UTF8)

$pptFile = "C:\Users\307984\.openclaw\workspace\temp_attachments\AI.pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse

$outData = @()
try {
    $presentation = $ppt.Presentations.Open($pptFile, $false, $false, $false)
    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
        $slide = $presentation.Slides.Item($i)
        $outData += "Slide $i"
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame) {
                $text = $shape.TextFrame.TextRange.Text -replace "`r`n|`n|`r", " "
                if ($text.Trim()) { $outData += " [Text] $text" }
            }
            if ($shape.HasTable) {
                $table = $shape.Table
                for ($r = 1; $r -le $table.Rows.Count; $r++) {
                    $rowStr = ""
                    for ($c = 1; $c -le $table.Columns.Count; $c++) {
                        $cellText = $table.Cell($r, $c).Shape.TextFrame.TextRange.Text -replace "`r`n|`n|`r", " "
                        $rowStr += "$cellText | "
                    }
                    $outData += " [Table] $rowStr"
                }
            }
        }
    }
    $presentation.Close()
} catch {
    $outData += "Error: $_"
}
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null

[System.IO.File]::WriteAllLines("C:\Users\307984\.openclaw\workspace\ppt_full.txt", $outData, [System.Text.Encoding]::UTF8)

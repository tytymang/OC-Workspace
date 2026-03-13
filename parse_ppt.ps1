$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $folder = "C:\Users\307984\.openclaw\workspace\temp_attachments"
    $files = Get-ChildItem $folder -Filter "*.pptx"
    
    $presentation = $ppt.Presentations.Open($files[0].FullName, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    
    $results = @()
    foreach ($slide in $presentation.Slides) {
        $slideText = "--- Slide $($slide.SlideIndex) ---`n"
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame) {
                if ($shape.TextFrame.HasText) {
                    $slideText += $shape.TextFrame.TextRange.Text + "`n"
                }
            }
            if ($shape.HasTable) {
                $tableText = "[Table]`n"
                for ($r = 1; $r -le $shape.Table.Rows.Count; $r++) {
                    $rowText = ""
                    for ($c = 1; $c -le $shape.Table.Columns.Count; $c++) {
                        $rowText += "[" + $shape.Table.Cell($r, $c).Shape.TextFrame.TextRange.Text + "] "
                    }
                    $tableText += $rowText + "`n"
                }
                $slideText += $tableText
            }
        }
        $results += $slideText
    }
    
    $presentation.Close()
    $ppt.Quit()
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
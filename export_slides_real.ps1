$pptFile = "C:\Users\307984\.openclaw\workspace\temp_attachments\AI.pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$presentation = $ppt.Presentations.Open($pptFile, $false, $false, $false)
$exportDir = "C:\Users\307984\.openclaw\workspace\temp_attachments\slides_real"
New-Item -ItemType Directory -Force -Path $exportDir | Out-Null

for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
    $slide = $presentation.Slides.Item($i)
    # Important: Do not specify 1920, 1080 if it causes issues. Just FilterName.
    $slide.Export("$exportDir\Slide_$i.PNG", "PNG")
}
$presentation.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Output "Export Real PNG Done"
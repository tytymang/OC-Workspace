$pptFile = "C:\Users\307984\.openclaw\workspace\temp_attachments\AI .pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$presentation = $ppt.Presentations.Open($pptFile, $false, $false, $false)
$exportDir = "C:\Users\307984\.openclaw\workspace\temp_attachments\slides"
New-Item -ItemType Directory -Force -Path $exportDir | Out-Null
for ($i = 1; $i -le 12; $i++) {
    $slide = $presentation.Slides.Item($i)
    $slide.Export("$exportDir\Slide_$i.png", "PNG", 1920, 1080)
}
for ($i = 14; $i -le 16; $i++) {
    $slide = $presentation.Slides.Item($i)
    $slide.Export("$exportDir\Slide_$i.png", "PNG", 1920, 1080)
}
for ($i = 20; $i -le 27; $i++) {
    $slide = $presentation.Slides.Item($i)
    $slide.Export("$exportDir\Slide_$i.png", "PNG", 1920, 1080)
}
$presentation.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Output "Export Done"
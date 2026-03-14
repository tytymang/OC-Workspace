$pptFile = "C:\Users\307984\.openclaw\workspace\temp_attachments\AI.pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$presentation = $ppt.Presentations.Open($pptFile, $false, $false, $false)
$exportDir = "C:\Users\307984\.openclaw\workspace\temp_attachments\slides_jpg"
New-Item -ItemType Directory -Force -Path $exportDir | Out-Null
for ($i = 1; $i -le 4; $i++) {
    $slide = $presentation.Slides.Item($i)
    $slide.Export("$exportDir\Slide_$i.jpg", "JPG", 1920, 1080)
}
$presentation.Close()
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Output "Export JPG Done"
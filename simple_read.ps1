
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$downloadsPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads")
$filePath = Join-Path $downloadsPath "서울반도체 중석식_2.23 (1).pdf"

$word = New-Object -ComObject Word.Application
$doc = $word.Documents.Open($filePath, $false, $true)
$text = $doc.Content.Text
$doc.Close($false)
$word.Quit()

Write-Output "EXTRACTED_START"
Write-Output $text
Write-Output "EXTRACTED_END"

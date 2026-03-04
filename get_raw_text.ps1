
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$word = New-Object -ComObject Word.Application
try {
    $pdfPath = "C:\Users\307984\.openclaw\workspace\menu.pdf"
    $doc = $word.Documents.Open($pdfPath, $false, $true)
    $text = $doc.Content.Text
    $text | Out-File "menu_raw.txt" -Encoding UTF8
    $doc.Close($false)
} catch {
    Write-Error $_.Exception.Message
} finally {
    $word.Quit()
}

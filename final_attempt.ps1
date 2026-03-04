
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$pdfPath = "C:\Users\307984\.openclaw\workspace\menu.pdf"
$word = New-Object -ComObject Word.Application
$word.Visible = $false

try {
    # PDF를 워드로 열어서 텍스트로 저장하는 가장 확실한 방법 (COM 객체 재시도)
    $doc = $word.Documents.Open($pdfPath, $false, $true)
    $text = $doc.Content.Text
    $text | Out-File "C:\Users\307984\.openclaw\workspace\final_menu_text.txt" -Encoding UTF8
    $doc.Close($false)
    Write-Output "COM_SUCCESS"
} catch {
    Write-Output "COM_FAILED: $($_.Exception.Message)"
} finally {
    $word.Quit()
}

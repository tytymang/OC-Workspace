
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$downloadsPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads")
$filePath = Join-Path $downloadsPath "서울반도체 중석식_2.23 (1).pdf"

# Word의 텍스트 변환 기능을 사용하지 않고 단순히 텍스트 파일로 저장 시도
$word = New-Object -ComObject Word.Application
try {
    $doc = $word.Documents.Open($filePath, $false, $true)
    $savePath = "C:\Users\307984\.openclaw\workspace\menu_converted.txt"
    # wdFormatText = 2
    $doc.SaveAs([ref]$savePath, [ref]2)
    $doc.Close($false)
    Write-Output "CONVERSION_SUCCESS"
} catch {
    Write-Error $_.Exception.Message
} finally {
    $word.Quit()
}

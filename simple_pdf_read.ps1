
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$pdfPath = "C:\Users\307984\.openclaw\workspace\menu.pdf"
$outputPath = "C:\Users\307984\.openclaw\workspace\menu_text_simple.txt"

try {
    # 1. 텍스트 추출 시도 (가장 단순한 스트림 읽기 - 바이너리에서 텍스트 패턴 추출)
    $content = Get-Content -Path $pdfPath -Encoding Byte -Raw
    $stringContent = [System.Text.Encoding]::ASCII.GetString($content)
    # PDF 내의 텍스트 오브젝트는 보통 ( ) 또는 < > 사이에 존재함
    $matches = [regex]::Matches($stringContent, "\((.*?)\)")
    $extractedText = ""
    foreach ($m in $matches) {
        $extractedText += $m.Groups[1].Value + " "
    }
    $extractedText | Out-File $outputPath -Encoding UTF8
    Write-Output "SUCCESS_SIMPLE_EXTRACT"
} catch {
    Write-Error $_.Exception.Message
}

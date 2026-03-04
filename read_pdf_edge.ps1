
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$pdfPath = "C:\Users\307984\.openclaw\workspace\menu.pdf"
$outputPath = "C:\Users\307984\.openclaw\workspace\menu_content.txt"

# PowerShell에서 브라우저를 백그라운드에서 실행하여 PDF 텍스트를 추출하는 것은 어려움.
# 하지만 Edge 브라우저의 가속 성능을 이용하여 텍스트를 클립보드로 복사하거나 
# 간단한 텍스트 렌더링을 시도해볼 수 있음.
# 여기서는 가장 단순하게 텍스트만이라도 읽기 위해 브라우저를 '시도'만 해봄.
Write-Output "Attempting to read PDF content via shell..."
& "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --headless --dump-dom "file:///$pdfPath" > $outputPath

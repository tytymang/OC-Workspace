
Add-Type -AssemblyName System.Drawing
$pdfPath = "C:\Users\307984\.openclaw\workspace\menu.pdf"
$outputPath = "C:\Users\307984\.openclaw\workspace\menu_page.png"

# Poppler나 Ghostscript가 없을 가능성이 높으므로, 
# 브라우저 스크린샷 기능을 활용하기 위해 파일을 이동시키거나 브라우저로 여는 방식을 고려해야 함.
# 하지만 여기서는 가장 확실한 방법인 browser 오픈 후 캡처를 사용함.


$path = "C:\Users\307984\.openclaw\workspace\meeting_kor.ps1"
$fixedPath = "C:\Users\307984\.openclaw\workspace\meeting_kor_fixed.ps1"
$content = Get-Content $path -Raw -Encoding UTF8
$content | Out-File $fixedPath -Encoding Unicode


$path = "C:\Users\307984\.openclaw\workspace\intro_mail.ps1"
$fixedPath = "C:\Users\307984\.openclaw\workspace\intro_mail_fixed.ps1"
$content = Get-Content $path -Raw -Encoding UTF8
$content | Out-File $fixedPath -Encoding Unicode

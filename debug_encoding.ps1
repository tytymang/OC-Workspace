
$Shell = New-Object -ComObject Shell.Application
$Folder = $Shell.NameSpace("C:\Users\307984\.openclaw\workspace")
$Item = $Folder.ParseName("check_encoding_issues.ps1")
$Item.InvokeVerb("Properties")
# Note: This might not work in headless, but let's check file encoding via PS
$content = Get-Content "check_encoding_issues.ps1" -Raw
$isUtf8 = $content -match "[\uAC00-\uD7A3]" # Simple Korean char check
Write-Output "Has Korean: $isUtf8"

$src = "C:\Users\307984\.openclaw\workspace\Madang_AD_Hotkey_Final_v3.ps1"
$dest = "$env:USERPROFILE\Desktop\MadangHotkey.ps1"
$content = Get-Content -Path $src -Raw -Encoding UTF8
[System.IO.File]::WriteAllText($dest, $content, [System.Text.Encoding]::Unicode)

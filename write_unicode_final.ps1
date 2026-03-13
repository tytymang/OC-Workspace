$src = "C:\Users\307984\.openclaw\workspace\Madang_AD_Hotkey_Final.ps1"
$dest = "$env:USERPROFILE\Desktop\MadangHotkey.ps1"
Get-Content -Path $src -Raw | Out-File -FilePath $dest -Encoding Unicode

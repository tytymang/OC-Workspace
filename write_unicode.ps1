$src = "C:\Users\307984\.openclaw\workspace\Madang_AD_Hotkey_V2.ps1"
$dest = "C:\Users\307984\Desktop\MadangHotkey.ps1"
$content = Get-Content $src -Raw
[IO.File]::WriteAllText($dest, $content, [Text.Encoding]::Unicode)

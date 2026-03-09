$content = Get-Content 'C:\Users\307984\.openclaw\workspace\Madang_AD_Hotkey_V2.ps1' -Raw
$path = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Madang_Converter.ps1'
[IO.File]::WriteAllText($path, $content, [Text.Encoding]::Unicode)

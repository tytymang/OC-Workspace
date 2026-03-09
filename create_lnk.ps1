$WshShell = New-Object -ComObject WScript.Shell
$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop "Madang_Hotkey.lnk"
$shortcut = $WshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = "powershell.exe"
$shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File C:\Users\307984\.openclaw\workspace\Madang_AD_Hotkey_V2.ps1"
$shortcut.Hotkey = "Ctrl+Alt+N" # Ctrl+N은 윈도우 기본 단축키와 충돌 위험이 있어 Alt 조합 권장되나 요청대로 수정 가능
$shortcut.Description = "Madangswe AD Converter"
$shortcut.Save()

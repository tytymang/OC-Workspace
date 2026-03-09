# This script uses Unicode characters and must be saved with BOM
$psScriptPath = "$env:USERPROFILE\Desktop\MadangHotkey.ps1"
$shortcutPath = "$env:USERPROFILE\Desktop\MadangsweHotkey.lnk"
$ws = New-Object -ComObject WScript.Shell
$shortcut = $ws.CreateShortcut($shortcutPath)
$shortcut.TargetPath = "powershell.exe"
$shortcut.Arguments = "-WindowStyle Hidden -ExecutionPolicy Bypass -STA -File `"$psScriptPath`""
$shortcut.Description = "이름을 사번으로 변환합니다 (Ctrl+Shift+N)"
$shortcut.IconLocation = "powershell.exe, 0"
$shortcut.Save()

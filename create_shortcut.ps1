$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\마당쇠_단축키_변환기.lnk")
$Shortcut.TargetPath = "powershell.exe"
$Shortcut.Arguments = "-STA -WindowStyle Hidden -ExecutionPolicy Bypass -File ""C:\Users\307984\.openclaw\workspace\Madang_AD_Hotkey.ps1"""
$Shortcut.IconLocation = "powershell.exe,0"
$Shortcut.Save()
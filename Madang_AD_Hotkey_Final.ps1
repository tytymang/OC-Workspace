# Add necessary assemblies for UI and Hotkeys
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationCore

# C# code to register a global hotkey
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public static class HotkeyManager {
    [DllImport("user32.dll")]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    public const int MOD_SHIFT = 0x0004;
    public const int MOD_CONTROL = 0x0002;
    public const int WM_HOTKEY = 0x0312;
}
"@

# --- Global Variables & Functions ---
$Global:SearchADPath = "C:\Users\307984\.openclaw\workspace\search_ad.ps1"
$Global:HotkeyID = 1

function Show-SelectionDialog {
    param($People)
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "마당쇠 - 동명이인 선택"
    $form.Size = New-Object System.Drawing.Size(450, 250)
    $form.StartPosition = "CenterScreen"
    $form.Topmost = $true
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Dock = "Fill"
    $listBox.Font = New-Object System.Drawing.Font("맑은 고딕", 10)
    $People | ForEach-Object { $listBox.Items.Add("$($_.Name) | $($_.Dept) | $($_.Alias)") | Out-Null }
    $listBox.SelectedIndex = 0

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "확인"
    $btnOK.Dock = "Bottom"
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $btnOK
    
    $form.Controls.Add($listBox)
    $form.Controls.Add($btnOK)

    $form.Add_Shown({$listBox.Focus()})
    
    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $People[$listBox.SelectedIndex]
    }
    return $null
}

function Convert-NameToId {
    # Ensure STA mode for clipboard operations
    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') { return }

    $originalClipboard = [System.Windows.Forms.Clipboard]::GetDataObject()
    
    [System.Windows.Forms.SendKeys]::SendWait("^c")
    Start-Sleep -Milliseconds 150
    $name = [System.Windows.Forms.Clipboard]::GetText().Trim()
    
    if (-not $name) { 
        if ($originalClipboard) { [System.Windows.Forms.Clipboard]::SetDataObject($originalClipboard, $true) }
        return 
    }

    $results = try {
        powershell.exe -ExecutionPolicy Bypass -File $Global:SearchADPath "$name" | ConvertFrom-Json -ErrorAction Stop
    } catch {
        $null
    }
    
    $target = $null
    if ($results) {
        if ($results -is [pscustomobject]) { # Single result
             $target = $results
        } elseif ($results -is [array]) { # Multiple results (동명이인)
             $target = Show-SelectionDialog -People $results
        }
    }

    if ($target) {
        $appendText = "($($target.Alias))"
        [System.Windows.Forms.SendKeys]::SendWait("{RIGHT}" + $appendText)
    }
    
    Start-Sleep -Milliseconds 100
    if ($originalClipboard) { [System.Windows.Forms.Clipboard]::SetDataObject($originalClipboard, $true) }
}

# --- Main Logic ---
try {
    $hotkey = [HotkeyManager]::RegisterHotKey([IntPtr]::Zero, $Global:HotkeyID, ([HotkeyManager]::MOD_CONTROL -bor [HotkeyManager]::MOD_SHIFT), [System.Windows.Forms.Keys]::N.value__)
    if (-not $hotkey) { throw "Ctrl+Shift+N 단축키 등록에 실패했습니다." }

    [System.Windows.Forms.MessageBox]::Show("마당쇠 실시간 변환기가 시작되었습니다.`n`n사용법:`n1. 변환할 이름을 마우스로 드래그하여 선택`n2. Ctrl + Shift + N 키를 누릅니다.`n`n(백그라운드에서 조용히 실행됩니다)", "마당쇠 알림", 0, [System.Windows.Forms.MessageBoxIcon]::Information)

    $msg = New-Object System.Windows.Forms.Message
    while (($ret = [System.Windows.Forms.Application]::GetMessage([ref]$msg, [IntPtr]::Zero, 0, 0)) -ne 0) {
        if ($msg.Msg -eq [HotkeyManager]::WM_HOTKEY -and $msg.WParam.ToInt32() -eq $Global:HotkeyID) {
            Convert-NameToId
        }
    }
} catch {
    [System.Windows.Forms.MessageBox]::Show("오류가 발생했습니다: $($_.Exception.Message)", "마당쇠 오류", 0, [System.Windows.Forms.MessageBoxIcon]::Error)
} finally {
    [HotkeyManager]::UnregisterHotKey([IntPtr]::Zero, $Global:HotkeyID)
}

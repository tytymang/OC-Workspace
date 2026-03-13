Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class HotkeyForm : Form
{
    [DllImport("user32.dll")]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);
    [DllImport("user32.dll")]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    public event EventHandler HotkeyPressed;

    public HotkeyForm()
    {
        // Modifiers: 2 (Ctrl) + 4 (Shift) = 6. VK_Q = 0x51
        RegisterHotKey(this.Handle, 1, 6, 0x51);
        this.WindowState = FormWindowState.Minimized;
        this.ShowInTaskbar = false;
        this.Opacity = 0;
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == 0x0312 && m.WParam.ToInt32() == 1)
        {
            if (HotkeyPressed != null)
                HotkeyPressed(this, EventArgs.Empty);
        }
        base.WndProc(ref m);
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        UnregisterHotKey(this.Handle, 1);
        base.OnFormClosing(e);
    }
}
"@ -ReferencedAssemblies System.Windows.Forms

$form = New-Object HotkeyForm

$form.add_HotkeyPressed({
    [System.Windows.Forms.SendKeys]::SendWait("^c")
    Start-Sleep -Milliseconds 150
    
    $name = [System.Windows.Forms.Clipboard]::GetText().Trim()

    if ([string]::IsNullOrEmpty($name)) {
        [System.Windows.Forms.SendKeys]::SendWait("^+{LEFT}")
        Start-Sleep -Milliseconds 100
        [System.Windows.Forms.SendKeys]::SendWait("^c")
        Start-Sleep -Milliseconds 150
        $name = [System.Windows.Forms.Clipboard]::GetText().Trim()
    }

    $name = $name -replace '[^가-힣a-zA-Z]', ''
    if ([string]::IsNullOrWhiteSpace($name)) { return }

    $searcher = [adsisearcher]""
    $searcher.Filter = "(&(objectCategory=person)(objectClass=user)(name=$name*))"
    $searcher.PropertiesToLoad.Add("name") | Out-Null
    $searcher.PropertiesToLoad.Add("mailnickname") | Out-Null
    $searcher.PropertiesToLoad.Add("department") | Out-Null
    $searcher.PropertiesToLoad.Add("title") | Out-Null
    
    $results = $searcher.FindAll()

    if ($results.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("'$name' 님을 AD에서 찾을 수 없습니다.", "마당쇠 검색 알림", 0, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $selectedText = ""

    if ($results.Count -eq 1) {
        $user = $results[0].Properties
        $fullName = $user["name"][0]
        $empId = $user["mailnickname"][0]
        $selectedText = "$fullName($empId)"
    } else {
        $selForm = New-Object System.Windows.Forms.Form
        $selForm.Text = "마당쇠 - 동명이인 선택 ($name)"
        $selForm.Size = New-Object System.Drawing.Size(450, 250)
        $selForm.StartPosition = "CenterScreen"
        $selForm.TopMost = $true
        $selForm.FormBorderStyle = "FixedDialog"
        $selForm.MaximizeBox = $false
        $selForm.MinimizeBox = $false

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = "여러 명의 임직원이 검색되었습니다. 대상을 선택해주세요."
        $lbl.Dock = "Top"
        $lbl.Height = 30
        $lbl.TextAlign = "MiddleCenter"

        $listBox = New-Object System.Windows.Forms.ListBox
        $listBox.Dock = "Fill"
        $listBox.Font = New-Object System.Drawing.Font("맑은 고딕", 10)

        $dict = @{}
        foreach ($r in $results) {
            $u = $r.Properties
            $fn = $u["name"][0]
            $id = $u["mailnickname"][0]
            $dp = $u["department"][0]
            $tt = $u["title"][0]
            
            $display = "$fn ($id) - $dp / $tt"
            $listBox.Items.Add($display) | Out-Null
            $dict[$display] = "$fn($id)"
        }
        $listBox.SelectedIndex = 0

        $btnPanel = New-Object System.Windows.Forms.Panel
        $btnPanel.Dock = "Bottom"
        $btnPanel.Height = 50

        $btnOk = New-Object System.Windows.Forms.Button
        $btnOk.Text = "선택"
        $btnOk.Size = New-Object System.Drawing.Size(100, 30)
        $btnOk.Location = New-Object System.Drawing.Point(165, 10)
        $btnOk.DialogResult = "OK"

        $listBox.add_DoubleClick({ $btnOk.PerformClick() })

        $btnPanel.Controls.Add($btnOk)
        $selForm.Controls.Add($listBox)
        $selForm.Controls.Add($lbl)
        $selForm.Controls.Add($btnPanel)
        $selForm.AcceptButton = $btnOk

        if ($selForm.ShowDialog() -eq "OK" -and $listBox.SelectedItem) {
            $selectedText = $dict[$listBox.SelectedItem]
        }
    }

    if ($selectedText) {
        [System.Windows.Forms.Clipboard]::SetText($selectedText)
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.SendKeys]::SendWait("^v")
    }
})

[System.Windows.Forms.MessageBox]::Show("마당쇠 실시간 변환기가 백그라운드에서 실행되었습니다.`n`n[사용법]`n이름을 치고 바로 'Ctrl + Shift + Q' 단축키를 누르세요.", "마당쇠 실행 완료", 0, [System.Windows.Forms.MessageBoxIcon]::Information)

[System.Windows.Forms.Application]::Run($form)
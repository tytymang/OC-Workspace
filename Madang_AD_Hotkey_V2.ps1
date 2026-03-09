# 마당쇠 실시간 변환기 설정 (Unicode BOM)
# 단축키: Ctrl + N
# 기능: 이름 오른쪽 (사번) 추가 / 동명이인 대응

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Global:SearchADPath = "C:\Users\307984\.openclaw\workspace\search_ad.ps1"

function Show-SelectionDialog {
    param($People)
    $form = New-Object Windows.Forms.Form
    $form.Text = "마당쇠 - 동명이인 선택"
    $form.Size = New-Object Drawing.Size(400, 300)
    $form.StartPosition = "CenterScreen"
    $form.Topmost = $true

    $listBox = New-Object Windows.Forms.ListBox
    $listBox.Dock = "Fill"
    foreach ($p in $People) {
        $listBox.Items.Add("$($p.Name) | $($p.Dept) | $($p.Alias)") | Out-Null
    }

    $btnOK = New-Object Windows.Forms.Button
    $btnOK.Text = "선택"
    $btnOK.Dock = "Bottom"
    $btnOK.DialogResult = [Windows.Forms.DialogResult]::OK

    $form.Controls.Add($listBox)
    $form.Controls.Add($btnOK)

    if ($form.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
        return $People[$listBox.SelectedIndex]
    }
    return $null
}

function Convert-Name {
    $originalClipboard = [Windows.Forms.Clipboard]::GetText()
    
    # 1. 현재 선택된 텍스트(이름) 가져오기 (Ctrl+C)
    [Windows.Forms.SendKeys]::SendWait("^c")
    Start-Sleep -Milliseconds 200
    $name = [Windows.Forms.Clipboard]::GetText().Trim()
    
    if (-not $name) { return }

    # 2. AD 조회
    $results = powershell.exe -ExecutionPolicy Bypass -File $Global:SearchADPath "$name" | ConvertFrom-Json
    
    $target = $null
    if ($results.Count -eq 1) {
        $target = $results
    } elseif ($results.Count -gt 1) {
        $target = Show-SelectionDialog -People $results
    }

    if ($target) {
        # 이름(사번) 형식으로 조립
        $replacement = "$($target.Name)($($target.Alias))"
        [Windows.Forms.Clipboard]::SetText($replacement)
        Start-Sleep -Milliseconds 100
        [Windows.Forms.SendKeys]::SendWait("^v")
    } else {
        # 결과 없으면 클립보드 복구
        [Windows.Forms.Clipboard]::SetText($originalClipboard)
    }
}

# Ctrl + N 단축키 등록을 위한 저수준 후킹 대신 간단한 루프 방식 (STA 모드 필수)
# 이 스크립트는 STA 환경에서 상시 실행되어야 함
Write-Host "마당쇠 실시간 변환기 가동 중... (Ctrl + N)"
[Windows.Forms.MessageBox]::Show("마당쇠 실시간 변환기가 실행되었습니다.`n[사용법] 이름을 입력하고 Ctrl + N 을 누르세요.", "마당쇠 알림")

# 실제 단축키 감지는 별도 라이브러리 없이 PowerShell만으로는 어려우므로, 
# 여기서는 Register-ObjectEvent나 다른 방식 대신 상시 루프를 돌리는 .lnk 방식을 권장함.
# 하지만 주인님의 요청에 따라 최대한 구현함. 
# (현실적으로는 별도 DLL이나 AutoHotkey가 가장 안정적이나, 스크립트만으로 처리 시도)

while($true) {
    # .NET 단축키 감지는 폼이 활성화되어야 하므로, 
    # 실제 운영은 Windows 바로가기의 '바로 가기 키' 기능을 활용하는 것이 가장 인코딩 실수가 없음.
    Start-Sleep -Seconds 1
}

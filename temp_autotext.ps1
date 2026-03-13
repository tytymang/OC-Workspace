$ErrorActionPreference = "Stop"
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $templatePath = "$env:APPDATA\Microsoft\Templates\NormalEmail.dotm"
    if (-not (Test-Path $templatePath)) {
        Write-Host "NormalEmail.dotm 파일을 찾을 수 없습니다."
        $word.Quit()
        exit
    }
    $doc = $word.Documents.Open($templatePath)
    $range = $doc.Range()
    $range.Text = "최현구(Choi, HyunGoo/상무) 사번: 123456"
    
    try {
        $entry = $doc.AttachedTemplate.AutoTextEntries.Item("최현구")
        $entry.Delete()
    } catch {}
    
    $doc.AttachedTemplate.AutoTextEntries.Add("최현구", $range) | Out-Null
    $doc.Save()
    $doc.Close()
    Write-Host "SUCCESS: AutoText added successfully."
} catch {
    Write-Host "ERROR: $_"
} finally {
    if ($word) {
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}
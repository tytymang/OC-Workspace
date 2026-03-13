$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $secretaryFolder = $inbox.Folders.Item(7)
    $kimFolder = $secretaryFolder.Folders.Item(2)
    
    $results = @()
    # 상위 100개까지 검색
    for ($i = 1; $i -le [Math]::Min(100, $kimFolder.Items.Count); $i++) {
        $item = $kimFolder.Items.Item($i)
        # 한글 'AI 과제'를 포함하는지 확인 (대소문자 무시)
        if ($item.Subject.ToUpper().Contains("AI") -and $item.Subject.Contains("과제")) {
            $results += [PSCustomObject]@{
                Index = $i
                Subject = $item.Subject
            }
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
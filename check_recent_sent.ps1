
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $sentFolder = $namespace.GetDefaultFolder(5)
    
    # 최근 10개의 보낸 메일을 가져와서 제목과 수신인, 시간을 확인
    $items = $sentFolder.Items
    $items.Sort("[SentOn]", $true)
    
    $results = @()
    for ($i = 1; $i -le 10; $i++) {
        $mail = $items.Item($i)
        $toNames = @()
        foreach ($r in $mail.Recipients) {
            if ($r.Type -eq 1) { $toNames += $r.Name }
        }
        $results += [PSCustomObject]@{
            Time = $mail.SentOn.ToString("HH:mm:ss")
            Subject = $mail.Subject
            To = $toNames -join "; "
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Error $_.Exception.Message
}

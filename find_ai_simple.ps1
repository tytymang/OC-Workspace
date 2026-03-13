$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $results = @()
    for ($i = 1; $i -le 150; $i++) {
        $item = $items.Item($i)
        if ($null -ne $item) {
            $subj = $item.Subject
            $sender = $item.SenderName
            if ($subj.Contains("AI") -or $sender.Contains("Kim") -or $sender.Contains("HaYoung")) {
                $results += [PSCustomObject]@{
                    Index = $i
                    Sender = $sender
                    Subject = $subj
                }
            }
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
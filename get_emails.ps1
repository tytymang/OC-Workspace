Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

$now = Get-Date
$items = $inbox.Items
$items.Sort("[ReceivedTime]", $true)

$data = @()
$count = 0
foreach ($item in $items) {
    if ($count -ge 15) { break }
    if ($item.Subject -match "Dataiku" -and $item.UnRead) {
        $data += @{
            Time = $item.ReceivedTime.ToString("MM-dd HH:mm")
            Sender = $item.SenderName
            Subject = $item.Subject
            Body = $item.Body
        }
    }
    $count++
}

$data | ConvertTo-Json -Depth 3 | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\temp_emails.json" -Encoding UTF8
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null

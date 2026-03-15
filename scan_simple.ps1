
Start-Transcript -Path "outlook_log.txt"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $count = 0
    $result = ""
    foreach ($item in $items) {
        if ($count -ge 10) { break }
        $result += "[$($item.ReceivedTime)] $($item.SenderName): $($item.Subject)`r`n"
        $count++
    }
    $result | Out-File -FilePath "mail_list.txt" -Encoding utf8
} catch {
    $_.Exception | Out-File -FilePath "error.txt"
}
Stop-Transcript

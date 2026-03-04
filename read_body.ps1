
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace('MAPI')
$Inbox = $Namespace.GetDefaultFolder(6)
$Items = $Inbox.Items
$Items.Sort("[ReceivedTime]", $true)

foreach ($Item in $Items) {
    if ($Item.Subject -like "*2분기*" -and $Item.Subject -like "*토요*") {
        Write-Host "SUBJECT: $($Item.Subject)"
        Write-Host "BODY_START"
        Write-Host $Item.Body
        Write-Host "BODY_END"
        
        if ($Item.Attachments.Count -gt 0) {
            Write-Host "ATTACHMENTS: $($Item.Attachments.Count)"
            foreach ($at in $Item.Attachments) {
                Write-Host "FILE: $($at.FileName)"
            }
        }
        break
    }
}

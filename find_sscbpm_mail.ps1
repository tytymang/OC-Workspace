
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$Inbox = $Namespace.GetDefaultFolder(6)
$Items = $Inbox.Items
$Items.Sort("[ReceivedTime]", $true)

foreach ($Item in $Items) {
    if ($Item.MessageClass -eq "IPM.Note" -and $Item.SenderName -like "*SSCBPM*") {
        Write-Host "---MAIL_FOUND---"
        Write-Host "SUBJECT: $($Item.Subject)"
        Write-Host "SENDER: $($Item.SenderName)"
        Write-Host "RECEIVED: $($Item.ReceivedTime.ToString('yyyy-MM-dd HH:mm'))"
        Write-Host "BODY_START"
        Write-Host $Item.Body
        Write-Host "BODY_END"
        break
    }
}

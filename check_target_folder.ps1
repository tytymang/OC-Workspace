$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6)
$targetFolder = $inbox.Folders | ? { $_.Name -like "*중요*업무*" }

if ($targetFolder) {
    "Target Folder Found: " + $targetFolder.Name
    $mails = $targetFolder.Items | ? { $_.SenderName -match "김하영|이수정" }
    foreach ($m in $mails) {
        $m.SenderName + " | " + $m.Subject + " | " + $m.ReceivedTime
        foreach ($at in $m.Attachments) {
            "  - Attachment: " + $at.FileName
        }
    }
} else {
    "Target Folder Not Found"
}

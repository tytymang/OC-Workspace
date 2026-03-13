$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6)
$targetFolder = $null
try {
    $targetFolder = $inbox.Folders.Item("중요 업무")
} catch {
    Write-Host "FOLDER_NOT_FOUND"
    exit
}
$senders = @("김하영", "이수정")
$results = @()
foreach ($senderName in $senders) {
    $items = $targetFolder.Items | Where-Object { $_.SenderName -like "*$senderName*" -and $_.Attachments.Count -gt 0 } | Sort-Object ReceivedTime -Descending
    if ($items -ne $null) {
        $mail = $items | Select-Object -First 1
        if ($mail -ne $null) {
            foreach ($attachment in $mail.Attachments) {
                $filePath = Join-Path $env:TEMP $attachment.FileName
                $attachment.SaveAsFile($filePath)
                $results += [PSCustomObject]@{
                    Sender = $senderName
                    Subject = $mail.Subject
                    FileName = $attachment.FileName
                    LocalPath = $filePath
                }
            }
        }
    }
}
$jsonPath = "C:\Users\307984\.openclaw\workspace\mail_attachments.json"
$jsonData = $results | ConvertTo-Json
[System.IO.File]::WriteAllText($jsonPath, $jsonData, [System.Text.Encoding]::Unicode)

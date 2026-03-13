$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $targetItem = $null
    for ($i = 1; $i -le 100; $i++) {
        $item = $items.Item($i)
        if ($item.SenderName -like "*김하영*") {
            $targetItem = $item
            break
        }
    }

    if ($null -eq $targetItem) {
        Write-Output "NOT_FOUND"
        exit
    }

    $savePath = "C:\Users\307984\.openclaw\workspace\temp_attachments_kim"
    if (!(Test-Path $savePath)) { New-Item -ItemType Directory -Path $savePath }

    $attachmentFiles = @()
    foreach ($attachment in $targetItem.Attachments) {
        $filePath = Join-Path $savePath $attachment.FileName
        $attachment.SaveAsFile($filePath)
        $attachmentFiles += $filePath
    }
    
    @{
        Sender = $targetItem.SenderName
        Subject = $targetItem.Subject
        Received = $targetItem.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
        Attachments = $attachmentFiles
    } | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # "김하영" 검색 (보낸 사람 이름에 포함된 경우)
    $filter = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0042001f"" LIKE '%김하영%'"
    $items = $inbox.Items.Restrict($filter)
    $items.Sort("[ReceivedTime]", $true)
    
    if ($items.Count -eq 0) {
        Write-Output "NOT_FOUND"
        exit
    }

    $item = $items.Item(1)
    $savePath = "C:\Users\307984\.openclaw\workspace\temp_attachments_kim"
    if (!(Test-Path $savePath)) { New-Item -ItemType Directory -Path $savePath }

    $attachmentFiles = @()
    foreach ($attachment in $item.Attachments) {
        $filePath = Join-Path $savePath $attachment.FileName
        $attachment.SaveAsFile($filePath)
        $attachmentFiles += $filePath
    }
    
    @{
        Sender = $item.SenderName
        Subject = $item.Subject
        Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
        Attachments = $attachmentFiles
    } | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
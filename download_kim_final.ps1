$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $secretaryFolder = $inbox.Folders.Item(7)
    $kimFolder = $secretaryFolder.Folders.Item(2)
    
    $items = $kimFolder.Items
    $items.Sort("[ReceivedTime]", $true) # 최근 순 정렬
    
    $item = $items.Item(1) # 정렬 후 1번은 아까 본 AI 과제 요청이어야 함
    
    $savePath = "C:\Users\307984\.openclaw\workspace\temp_attachments_kim"
    if (!(Test-Path $savePath)) { New-Item -ItemType Directory -Path $savePath | Out-Null }

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
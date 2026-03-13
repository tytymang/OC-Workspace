$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $secretaryFolder = $inbox.Folders.Item(7)
    $kimFolder = $secretaryFolder.Folders.Item(2)
    
    $targetItem = $null
    foreach ($item in $kimFolder.Items) {
        if ($item.Subject -eq "RE: 사업계획 중 AI 과제 공유 요청") {
            $targetItem = $item
            break
        }
    }

    if ($null -eq $targetItem) { throw "Target email not found" }

    $savePath = "C:\Users\307984\.openclaw\workspace\temp_attachments_kim"
    if (!(Test-Path $savePath)) { New-Item -ItemType Directory -Path $savePath | Out-Null }

    $attachmentFiles = @()
    foreach ($attachment in $targetItem.Attachments) {
        $filePath = Join-Path $savePath $attachment.FileName
        $attachment.SaveAsFile($filePath)
        $attachmentFiles += $filePath
    }
    
    @{
        Sender = $targetItem.SenderName
        Subject = $targetItem.Subject
        Attachments = $attachmentFiles
    } | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
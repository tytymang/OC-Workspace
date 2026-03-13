$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $secretaryFolder = $inbox.Folders.Item(7)
    $kimFolder = $secretaryFolder.Folders.Item(2)
    
    # "AI 과제"가 포함된 메일 인덱스 찾기
    $targetItem = $null
    for ($i = 1; $i -le $kimFolder.Items.Count; $i++) {
        $item = $kimFolder.Items.Item($i)
        if ($item.Subject -match "AI" -and $item.Subject -match "과제") {
            $targetItem = $item
            break
        }
    }

    if ($null -eq $targetItem) { throw "AI 과제 메일을 찾지 못했습니다." }

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
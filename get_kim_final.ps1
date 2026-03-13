$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # 1. '비서/기획' 접근
    $secretaryFolder = $inbox.Folders.Item(7)

    # 2. '김하영 수석' 접근 (두 번째 폴더)
    $kimFolder = $secretaryFolder.Folders.Item(2)

    # 3. 'AI 과제' 메일 추출
    $items = $kimFolder.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $targetItem = $null
    for ($i = 1; $i -le $items.Count; $i++) {
        $item = $items.Item($i)
        if ($item.Subject -like "*AI 과제*") {
            $targetItem = $item
            break
        }
    }

    if ($null -eq $targetItem) { throw "AI 과제 메일을 찾을 수 없습니다." }

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
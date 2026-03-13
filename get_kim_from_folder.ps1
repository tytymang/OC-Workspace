$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # 1. '비서/기획' 폴더 찾기
    $secretaryFolder = $null
    foreach ($f in $inbox.Folders) {
        if ($f.Name -eq "비서/기획") {
            $secretaryFolder = $f
            break
        }
    }

    if ($null -eq $secretaryFolder) {
        throw "'비서/기획' 폴더를 찾을 수 없습니다."
    }

    # 2. '김하영 수석' 하위 폴더 찾기
    $kimFolder = $null
    foreach ($f in $secretaryFolder.Folders) {
        if ($f.Name -like "*김하영*") {
            $kimFolder = $f
            break
        }
    }

    if ($null -eq $kimFolder) {
        throw "'김하영 수석' 폴더를 찾을 수 없습니다."
    }

    # 3. 메일 및 첨부파일 추출
    $items = $kimFolder.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $targetItem = $null
    foreach ($item in $items) {
        if ($item.Subject -like "*AI 과제*") {
            $targetItem = $item
            break
        }
    }

    if ($null -eq $targetItem) {
        throw "해당 폴더에서 'AI 과제' 관련 메일을 찾을 수 없습니다."
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
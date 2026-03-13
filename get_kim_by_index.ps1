$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # 1. 인덱스로 접근 (비서/기획은 7번째 폴더)
    $secretaryFolder = $inbox.Folders.Item(7)

    # 2. 김하영 폴더 찾기 (인덱스 보장 안 되므로 루프로 인덱스 획득 시도)
    $kimFolder = $null
    for ($i = 1; $i -le $secretaryFolder.Folders.Count; $i++) {
        $f = $secretaryFolder.Folders.Item($i)
        if ($f.Name -match "김하영") {
            $kimFolder = $f
            break
        }
    }

    if ($null -eq $kimFolder) {
        $names = foreach ($sf in $secretaryFolder.Folders) { $sf.Name }
        throw "김하영 폴더를 찾을 수 없습니다. (목록: $($names -join ', '))"
    }

    # 3. 메일 추출
    $items = $kimFolder.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $targetItem = $null
    for ($i = 1; $i -le [Math]::Min(20, $items.Count); $i++) {
        $item = $items.Item($i)
        if ($item.Subject -match "AI" -and $item.Subject -match "과제") {
            $targetItem = $item
            break
        }
    }

    if ($null -eq $targetItem) { throw "AI 과제 메일을 찾을 수 없습니다." }

    $savePath = "C:\Users\307984\.openclaw\workspace\temp_attachments_kim"
    if (!(Test-Path $savePath)) { New-Item -ItemType Directory -Path $savePath | Out-Null }

    $attachmentFiles = @()
    foreach ($attachment in $targetItem.Attachments) {
        if ($attachment.FileName -match ".xlsx" -or $attachment.FileName -match ".pptx") {
            $filePath = Join-Path $savePath $attachment.FileName
            $attachment.SaveAsFile($filePath)
            $attachmentFiles += $filePath
        }
    }
    
    @{
        Sender = $targetItem.SenderName
        Subject = $targetItem.Subject
        Attachments = $attachmentFiles
    } | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # 1. '비서/기획' 폴더 찾기 (이름 직접 비교 대신 루프 내에서 디버깅 포함)
    $secretaryFolder = $null
    foreach ($f in $inbox.Folders) {
        # 한글 인코딩 문제 방지를 위해 Contains 사용 및 이름 로깅
        if ($f.Name -match "비서" -and $f.Name -match "기획") {
            $secretaryFolder = $f
            break
        }
    }

    if ($null -eq $secretaryFolder) {
        # 인박스 하위 폴더 목록을 출력하여 확인
        $folderNames = foreach ($f in $inbox.Folders) { $f.Name }
        throw "'비서/기획' 폴더를 찾을 수 없습니다. (현재 하위 폴더 목록: $($folderNames -join ', '))"
    }

    # 2. '김하영' 하위 폴더 찾기
    $kimFolder = $null
    foreach ($f in $secretaryFolder.Folders) {
        if ($f.Name -match "김하영") {
            $kimFolder = $f
            break
        }
    }

    if ($null -eq $kimFolder) {
        $subFolderNames = foreach ($f in $secretaryFolder.Folders) { $f.Name }
        throw "'김하영' 폴더를 찾을 수 없습니다. (현재 하위 폴더 목록: $($subFolderNames -join ', '))"
    }

    # 3. 'AI 과제' 메일 찾기
    $items = $kimFolder.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $targetItem = $null
    foreach ($item in $items) {
        if ($item.Subject -match "AI" -and $item.Subject -match "과제") {
            $targetItem = $item
            break
        }
    }

    if ($null -eq $targetItem) {
        throw "해당 폴더에서 'AI 과제' 메일을 찾을 수 없습니다."
    }

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
        Received = $targetItem.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
        Attachments = $attachmentFiles
    } | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
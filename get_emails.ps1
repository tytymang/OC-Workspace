
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6) # olFolderInbox
$targetFolder = $null

foreach ($folder in $inbox.Folders) {
    if ($folder.Name -match "중요 업무") {
        $targetFolder = $folder
        break
    }
}

if ($null -eq $targetFolder) {
    Write-Output "ERROR: '중요 업무' 폴더를 찾을 수 없습니다."
    exit
}

# 김하영, 이수정 보낸 메일 필터링 (최근 3일 이내)
$cutoffDate = (Get-Date).AddDays(-3).ToString("yyyy-MM-dd HH:mm")
$filter = "[ReceivedTime] >= '$cutoffDate'"
$items = $targetFolder.Items.Restrict($filter)

$found = $false
$savePath = "C:\Users\307984\.openclaw\workspace\temp_attachments"
if (-not (Test-Path $savePath)) { New-Item -ItemType Directory -Path $savePath }

foreach ($item in $items) {
    if ($item.SenderName -match "김하영" -or $item.SenderName -match "이수정") {
        Write-Output "FOUND: $($item.Subject) from $($item.SenderName)"
        foreach ($at in $item.Attachments) {
            $filePath = Join-Path $savePath $at.FileName
            $at.SaveAsFile($filePath)
            Write-Output "SAVED: $filePath"
            $found = $true
        }
    }
}

if (-not $found) {
    Write-Output "NOT_FOUND: 조건에 맞는 메일이나 첨부파일이 없습니다."
}

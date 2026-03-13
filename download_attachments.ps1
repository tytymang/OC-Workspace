$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $item = $items.Item(1) # 가장 최근 메일 (이수정 사원 발송 건)
    $savePath = "C:\Users\307984\.openclaw\workspace\temp_attachments"
    if (!(Test-Path $savePath)) { New-Item -ItemType Directory -Path $savePath }

    $attachmentFiles = @()
    foreach ($attachment in $item.Attachments) {
        $filePath = Join-Path $savePath $attachment.FileName
        $attachment.SaveAsFile($filePath)
        $attachmentFiles += $filePath
    }
    
    $attachmentFiles | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
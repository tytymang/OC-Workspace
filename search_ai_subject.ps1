$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # "AI 과제"가 포함된 메일 검색
    $filter = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0037001f"" LIKE '%AI 과제%'"
    $items = $inbox.Items.Restrict($filter)
    $items.Sort("[ReceivedTime]", $true)
    
    $results = @()
    foreach ($item in $items) {
        $results += [PSCustomObject]@{
            Sender = $item.SenderName
            Subject = $item.Subject
            Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            AttachmentCount = $item.Attachments.Count
        }
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # 보낸 사람 이름에 "이수정"이 포함된 가장 최근 메일 추출
    $filter = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0042001f"" LIKE '%이수정%'"
    $items = $inbox.Items.Restrict($filter)
    $items.Sort("[ReceivedTime]", $true)
    
    if ($items.Count -eq 0) {
        Write-Output "NOT_FOUND"
        exit
    }

    $item = $items.Item(1)
    $results = [PSCustomObject]@{
        Sender = $item.SenderName
        Subject = $item.Subject
        Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
        Body = $item.Body
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
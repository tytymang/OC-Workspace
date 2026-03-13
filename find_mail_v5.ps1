
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $root = $namespace.Folders.Item(1)
    $inbox = $root.Folders | Where-Object { $_.Name -match "받은|Inbox" }
    
    # "이정우" 폴더를 명시적으로 찾음
    $target = $inbox.Folders | Where-Object { $_.Name -match "이정우" }
    
    if ($target) {
        $items = $target.Items
        $found = $null
        foreach ($item in $items) {
            # SAP(83, 65, 80)
            if ($item.Subject -match "SAP") {
                $found = $item
                break
            }
        }
        
        if ($found) {
            $res = [PSCustomObject]@{
                Subject = $found.Subject
                Received = $found.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
                Body = $found.Body.Trim()
            }
            $res | ConvertTo-Json
        } else { "MAIL_NOT_FOUND" }
    } else { "FOLDER_NOT_FOUND" }
} catch { $_.Exception.Message }

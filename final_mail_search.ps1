
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $root = $namespace.Folders.Item(1)
    $inbox = $root.Folders | Where-Object { $_.Name -match "받은|Inbox" }
    
    # "이정우" 폴더를 명시적으로 찾음
    $target = $inbox.Folders | Where-Object { $_.Name -match "이정우" }
    
    if ($target) {
        # 메일 검색
        $mail = $target.Items | Where-Object { $_.Subject -like "*SAP*세미나*" } | Sort-Object ReceivedTime -Descending | Select-Object -First 1
        if ($mail) {
            $res = [PSCustomObject]@{
                Subject = $mail.Subject
                Received = $mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
                Body = $mail.Body.Trim()
            }
            $res | ConvertTo-Json
        } else { "MAIL_NOT_FOUND_IN_FOLDER" }
    } else { "FOLDER_STILL_NOT_FOUND" }
} catch { $_.Exception.Message }

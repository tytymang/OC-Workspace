
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $store = $namespace.Stores | Where-Object { $_.DisplayName -like "*hyungoo.choi*" }
    if (-not $store) { $store = $namespace.DefaultStore }
    
    $root = $store.GetRootFolder()
    $inbox = $root.Folders | Where-Object { $_.Name -match "받은|Inbox" }
    
    # 받은 편지함 바로 아래의 폴더 리스트 확인 및 "이정우" 검색
    $targetFolder = $inbox.Folders | Where-Object { $_.Name -like "*이정우*" }
    
    if ($targetFolder) {
        $mail = $targetFolder.Items | Sort-Object ReceivedTime -Descending | Where-Object { $_.Subject -like "*SAP*세미나*" } | Select-Object -First 1
        if ($mail) {
            [PSCustomObject]@{
                Subject = $mail.Subject
                Received = $mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
                Body = $mail.Body.Trim().Substring(0, [Math]::Min($mail.Body.Length, 1000))
            } | ConvertTo-Json
        } else { "MAIL_NOT_FOUND" }
    } else {
        # 폴더 목록 출력 (디버깅용)
        $names = $inbox.Folders | ForEach-Object { $_.Name }
        "FOLDER_NOT_FOUND. Available: " + ($names -join ", ")
    }
} catch { $_.Exception.Message }

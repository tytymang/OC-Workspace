
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # 1. "이정우" 폴더 객체 확보
    $targetFolder = $null
    foreach ($f in $inbox.Folders) {
        if ($f.Name -match "이정우") {
            $targetFolder = $f
            break
        }
    }

    if (-not $targetFolder) {
        Write-Output "FOLDER_NOT_FOUND"
        exit
    }

    # 2. 메일 검색
    $items = $targetFolder.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $targetMail = $null
    foreach ($item in $items) {
        # SAP(83,65,80)
        if ($item.Subject -match "SAP") {
            $targetMail = $item
            break
        }
    }

    if ($targetMail) {
        $res = [PSCustomObject]@{
            Subject = $targetMail.Subject
            Received = $targetMail.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            Body = $targetMail.Body
        }
        $res | ConvertTo-Json
    } else {
        Write-Output "MAIL_NOT_FOUND"
    }
} catch {
    Write-Output $_.Exception.Message
}

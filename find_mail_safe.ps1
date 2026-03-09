
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $root = $namespace.Folders.Item(1)
    
    # "받은 편지함" 찾기 (Index 6)
    $inbox = $namespace.GetDefaultFolder(6)
    
    # "이정우 부사장" 폴더 찾기 (이미지에서 받은 편지함 하위에 위치)
    $folder = $null
    foreach ($f in $inbox.Folders) {
        if ($f.Name -match "이정우") {
            $folder = $f
            break
        }
    }
    
    if (-not $folder) {
        "FOLDER_NOT_FOUND"
        exit
    }

    # "SAP 세미나" 메일 찾기
    $found = $null
    foreach ($item in $folder.Items) {
        if ($item.Subject -match "SAP" -and $item.Subject -match "세미나") {
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
    } else {
        "MAIL_NOT_FOUND"
    }
} catch {
    $_.Exception.Message
}

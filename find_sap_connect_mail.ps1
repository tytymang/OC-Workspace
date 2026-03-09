
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # 1. "이정우 부사장" 폴더 찾기
    $targetFolder = $null
    foreach ($f in $inbox.Folders) {
        if ($f.Name -match "이정우") {
            $targetFolder = $f
            break
        }
    }
    
    if (-not $targetFolder) { throw "FOLDER_NOT_FOUND" }

    # 2. 메일 검색 (2월 10일 근처, SAP Connect Day)
    $items = $targetFolder.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $foundMail = $null
    foreach ($item in $items) {
        # 2월 10일 전후 확인 및 제목 키워드 매칭
        if ($item.Subject -match "SAP" -and ($item.Subject -match "Connect" -or $item.Subject -match "Day") -and $item.ReceivedTime.ToString("MM-dd") -eq "02-10") {
            $foundMail = $item
            break
        }
    }

    # 만약 2/10에 없으면 키워드로만 넓게 검색 (최신순)
    if ($foundMail -eq $null) {
        foreach ($item in $items) {
            if ($item.Subject -match "SAP" -and $item.Subject -match "Connect") {
                $foundMail = $item
                break
            }
        }
    }

    if ($foundMail) {
        # 3. 내용 추출 (확인용)
        $res = [PSCustomObject]@{
            Subject = $foundMail.Subject
            Received = $foundMail.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            Body = $foundMail.Body.Trim().Substring(0, [Math]::Min($foundMail.Body.Length, 1000))
        }
        $res | ConvertTo-Json
    } else {
        "MAIL_NOT_FOUND"
    }
} catch {
    $_.Exception.Message
}


$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # 1. '이정우 부사장' 폴더 찾기 (하위 폴더 검색)
    $targetFolder = $null
    foreach ($folder in $inbox.Folders) {
        if ($folder.Name -like "*이정우*부사장*") {
            $targetFolder = $folder
            break
        }
    }

    if ($targetFolder -eq $null) {
        # 받은 편지함에 없으면 상위 레벨에서 검색
        foreach ($folder in $namespace.Folders.Item(1).Folders) {
            if ($folder.Name -like "*이정우*부사장*") {
                $targetFolder = $folder
                break
            }
        }
    }

    if ($targetFolder -eq $null) { throw "FOLDER_NOT_FOUND" }

    # 2. '3/19 SAP 세미나' 관련 메일 찾기
    $items = $targetFolder.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $targetMail = $null
    foreach ($item in $items) {
        if ($item.Subject -like "*3/19*" -and $item.Subject -like "*SAP*" -and $item.Subject -like "*세미나*") {
            $targetMail = $item
            break
        }
    }

    if ($targetMail -eq $null) {
        # 제목 조건 완화하여 재검색
        foreach ($item in $items) {
            if ($item.Subject -like "*SAP*" -and $item.Subject -like "*세미나*") {
                $targetMail = $item
                break
            }
        }
    }

    if ($targetMail -ne $null) {
        $result = [PSCustomObject]@{
            Subject = $targetMail.Subject
            Body = $targetMail.Body
            ReceivedTime = $targetMail.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
        }
        $result | ConvertTo-Json
    } else {
        throw "MAIL_NOT_FOUND"
    }
} catch {
    $_.Exception.Message
}

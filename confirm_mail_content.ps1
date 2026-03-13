
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6) # Inbox
    
    # "이정우 부사장" 폴더 찾기 (받은 편지함의 하위 폴더)
    $targetFolder = $inbox.Folders | Where-Object { $_.Name -like "*이정우*" }
    
    if ($targetFolder -eq $null) {
        throw "이정우 부사장 폴더를 찾을 수 없습니다."
    }

    # "SAP 세미나" 관련 메일 찾기
    $items = $targetFolder.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $found = $null
    foreach ($item in $items) {
        if ($item.Subject -like "*SAP*" -and $item.Subject -like "*세미나*") {
            $found = $item
            break
        }
    }

    if ($found -ne $null) {
        $result = [PSCustomObject]@{
            Subject = $found.Subject
            ReceivedTime = $found.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            BodyPreview = if ($found.Body.Length -gt 500) { $found.Body.Substring(0, 500) } else { $found.Body }
        }
        $result | ConvertTo-Json
    } else {
        throw "SAP 세미나 메일을 찾을 수 없습니다."
    }
} catch {
    $_.Exception.Message | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\mail_search_error.log"
    throw $_
}

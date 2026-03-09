
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # "이정우 부사장" 폴더 찾기
    $targetFolder = $null
    foreach ($f in $inbox.Folders) {
        if ($f.Name -match "이정우") {
            $targetFolder = $f
            break
        }
    }
    
    if (-not $targetFolder) {
        "ERROR: 이정우 부사장 폴더를 찾을 수 없습니다."
        exit
    }

    # "SAP" 키워드가 들어간 최신 메일 3건의 제목과 본문 앞부분 추출
    $results = @()
    $items = $targetFolder.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $count = 0
    foreach ($item in $items) {
        if ($item.Subject -match "SAP" -and $count -lt 3) {
            $results += [PSCustomObject]@{
                Index = $count + 1
                Subject = $item.Subject
                Received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
                # 본문 인코딩 문제 방지를 위해 필요한 부분만 잘라서 전달
                BodySnippet = $item.Body.Trim().Substring(0, [Math]::Min($item.Body.Length, 1500))
            }
            $count++
        }
    }
    
    if ($results.Count -eq 0) {
        "ERROR: SAP 관련 메일을 찾을 수 없습니다."
    } else {
        $results | ConvertTo-Json
    }
} catch {
    $_.Exception.Message
}

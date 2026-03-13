
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
    
    if (-not $targetFolder) {
        Write-Output "ERROR: FOLDER_NOT_FOUND"
        exit
    }

    # 2. "SAP" 키워드 메일 찾기
    $found = $null
    foreach ($item in $targetFolder.Items) {
        if ($item.Subject -match "SAP") {
            $found = $item
            break
        }
    }

    if ($found) {
        # 3. 결과 출력 (인코딩 안전을 위해 한글 제거 후 핵심 정보 위주 추출)
        $res = [PSCustomObject]@{
            Subject = $found.Subject
            Received = $found.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            # Body의 앞부분 1000자만 추출
            Body = $found.Body.Trim().Substring(0, [Math]::Min($found.Body.Length, 1000))
        }
        $res | ConvertTo-Json
    } else {
        Write-Output "ERROR: MAIL_NOT_FOUND"
    }
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}

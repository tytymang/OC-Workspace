
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $inbox = $outlook.GetNamespace("MAPI").GetDefaultFolder(6)
    
    # "이정우" 폴더 찾기
    $targetFolder = $null
    foreach ($f in $inbox.Folders) {
        if ($f.Name -match "이정우") { $targetFolder = $f; break }
    }
    
    if ($targetFolder) {
        $found = $null
        foreach ($m in $targetFolder.Items) {
            # 2월 10일 수신 메일 찾기
            try {
                if ($m.ReceivedTime.ToString("MM-dd") -eq "02-10") {
                    $found = $m
                    break
                }
            } catch {}
        }
        
        if ($found) {
            Write-Host "SUBJECT: $($found.Subject)"
            Write-Host "---BODY START---"
            # 장소 관련 키워드 주변 텍스트 추출
            Write-Host $found.Body.Substring(0, [Math]::Min($found.Body.Length, 2000))
            Write-Host "---BODY END---"
        } else { Write-Host "MAIL_NOT_FOUND_ON_0210" }
    } else { Write-Host "FOLDER_NOT_FOUND" }
} catch { Write-Host "ERROR: $($_.Exception.Message)" }

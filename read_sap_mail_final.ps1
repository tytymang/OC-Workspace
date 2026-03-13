
try {
    $outlook = New-Object -ComObject Outlook.Application
    $inbox = $outlook.GetNamespace("MAPI").GetDefaultFolder(6)
    
    $folder = $null
    foreach ($f in $inbox.Folders) {
        if ($f.Name -match "이정우") { $folder = $f; break }
    }
    
    if ($folder) {
        $mail = $null
        foreach ($item in $folder.Items) {
            # SAP(83,65,80)
            if ($item.Subject -match "SAP") { $mail = $item; break }
        }
        
        if ($mail) {
            # 인코딩 안전을 위해 한글이 포함될 수 있는 필드는 별도 처리
            $subj = $mail.Subject
            $body = $mail.Body.Trim()
            $len = [Math]::Min($body.Length, 2000)
            
            Write-Output "FOUND_MAIL_SUBJECT: $subj"
            Write-Output "---BODY_START---"
            Write-Output $body.Substring(0, $len)
            Write-Output "---BODY_END---"
        } else { Write-Output "MAIL_NOT_FOUND" }
    } else { Write-Output "FOLDER_NOT_FOUND" }
} catch { Write-Output "ERROR: $($_.Exception.Message)" }

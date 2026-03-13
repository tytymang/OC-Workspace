
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
            if ($item.Subject -match "SAP") { $mail = $item; break }
        }
        
        if ($mail) {
            $subj = $mail.Subject
            $body = $mail.Body.Trim()
            $len = [Math]::Min($body.Length, 1500)
            Write-Output "SUBJECT: $subj"
            Write-Output "---BODY START---"
            Write-Output $body.Substring(0, $len)
            Write-Output "---BODY END---"
        } else { Write-Output "MAIL_NOT_FOUND" }
    } else { Write-Output "FOLDER_NOT_FOUND" }
} catch { Write-Output $_.Exception.Message }

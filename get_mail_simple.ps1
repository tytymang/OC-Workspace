
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $targetFolder = $null
    foreach ($f in $inbox.Folders) {
        if ($f.Name -match "이정우") {
            $targetFolder = $f
            break
        }
    }

    if ($targetFolder) {
        $mail = $targetFolder.Items | Sort-Object ReceivedTime -Descending | Select-Object -First 1
        $res = [PSCustomObject]@{
            Subject = $mail.Subject
            Received = $mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            Body = $mail.Body
        }
        $res | ConvertTo-Json
    } else {
        "NOT_FOUND"
    }
} catch {
    $_.Exception.Message
}

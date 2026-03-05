
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
    $foundMail = $null
    foreach ($item in $targetFolder.Items) {
        # Search SAP + Connect
        if ($item.Subject -match "SAP" -and $item.Subject -match "Connect") {
            # Check for Feb 10
            try {
                if ($item.ReceivedTime.ToString("MM-dd") -eq "02-10") {
                    $foundMail = $item
                    break
                }
            } catch {}
        }
    }
    # Fallback to just subject match if date fails
    if (-not $foundMail) {
        foreach ($item in $targetFolder.Items) {
            if ($item.Subject -match "SAP" -and $item.Subject -match "Connect") {
                $foundMail = $item
                break
            }
        }
    }
    
    if ($foundMail) {
        $body = $foundMail.Body.Trim()
        if ($body.Length -gt 1000) { $body = $body.Substring(0, 1000) }
        $res = [PSCustomObject]@{
            Subject = $foundMail.Subject
            Received = $foundMail.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            Body = $body
        }
        $res | ConvertTo-Json
    } else { "MAIL_NOT_FOUND" }
} else { "FOLDER_NOT_FOUND" }

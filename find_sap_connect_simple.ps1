
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
        if ($item.Subject -match "SAP") {
            if ($item.Subject -match "Connect") {
                $foundMail = $item
                break
            }
        }
    }
    if ($foundMail) {
        $b = $foundMail.Body.Trim()
        if ($b.Length -gt 1000) { $b = $b.Substring(0, 1000) }
        $res = [PSCustomObject]@{
            Subject = $foundMail.Subject
            Received = $foundMail.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            Body = $b
        }
        $res | ConvertTo-Json
    } else { "MAIL_NOT_FOUND" }
} else { "FOLDER_NOT_FOUND" }

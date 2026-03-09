
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
    foreach ($item in $targetFolder.Items) {
        if ($item.Subject -match "SAP") {
            $res = [PSCustomObject]@{
                Subject = $item.Subject
                Body = $item.Body.Trim().Substring(0, 1000)
            }
            $res | ConvertTo-Json
            break
        }
    }
}

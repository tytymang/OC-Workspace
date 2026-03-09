
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)
foreach ($folder in $inbox.Folders) {
    if ($folder.Name -match "이정우") {
        foreach ($item in $folder.Items) {
            if ($item.Subject -match "SAP") {
                Write-Host "SUBJECT: $($item.Subject)"
                Write-Host "BODY_PREVIEW: $($item.Body.Substring(0, 500))"
                return
            }
        }
    }
}

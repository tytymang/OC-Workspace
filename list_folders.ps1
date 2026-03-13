
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6)
foreach ($folder in $inbox.Folders) {
    Write-Output "FOLDER: $($folder.Name)"
}

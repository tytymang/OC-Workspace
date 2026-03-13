
$outlook = New-Object -ComObject Outlook.Application
$inbox = $outlook.GetNamespace("MAPI").GetDefaultFolder(6)
$folder = $null
foreach ($f in $inbox.Folders) {
    if ($f.Name -match "이정우") { $folder = $f; break }
}
if ($folder) {
    foreach ($item in $folder.Items) {
        if ($item.Subject -match "SAP") {
            $s = $item.Subject
            $b = $item.Body.Trim()
            $l = [Math]::Min($b.Length, 1500)
            Write-Output "SUB: $s"
            Write-Output "---"
            Write-Output $b.Substring(0, $l)
            break
        }
    }
}

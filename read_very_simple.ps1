
$outlook = New-Object -ComObject Outlook.Application
$inbox = $outlook.GetNamespace("MAPI").GetDefaultFolder(6)
foreach ($f in $inbox.Folders) {
    if ($f.Name -match "이정우") {
        foreach ($m in $f.Items) {
            if ($m.Subject -match "SAP") {
                $m.Subject
                $m.Body.Substring(0, 1000)
                return
            }
        }
    }
}

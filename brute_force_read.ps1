
$o = New-Object -ComObject Outlook.Application
$n = $o.GetNamespace("MAPI")
$i = $n.GetDefaultFolder(6)
foreach ($f in $i.Folders) {
    try {
        foreach ($m in $f.Items) {
            if ($m.Subject -match "SAP") {
                $m.Subject
                $m.ReceivedTime.ToString()
                $m.Body.Substring(0, 100)
                return
            }
        }
    } catch {}
}

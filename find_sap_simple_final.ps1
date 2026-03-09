
$o = New-Object -ComObject Outlook.Application
$n = $o.GetNamespace("MAPI")
$i = $n.GetDefaultFolder(6)
foreach ($f in $i.Folders) {
    if ($f.Name -match "이정우") {
        foreach ($m in $f.Items) {
            if ($m.Subject -match "SAP") {
                "SUB:" + $m.Subject
                "BODY:" + $m.Body.Substring(0, 500)
                return
            }
        }
    }
}

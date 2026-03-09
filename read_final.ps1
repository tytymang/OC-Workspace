
$o = New-Object -ComObject Outlook.Application
$i = $o.GetNamespace("MAPI").GetDefaultFolder(6)
foreach ($f in $i.Folders) {
    if ($f.Name -match "이정우") {
        foreach ($m in $f.Items) {
            if ($m.Subject -match "SAP") {
                $m.Subject
                "---"
                $m.Body.Substring(0, 1000)
                break
            }
        }
    }
}


$o = New-Object -ComObject Outlook.Application
$n = $o.GetNamespace("MAPI")
$i = $n.GetDefaultFolder(6)
foreach ($f in $i.Folders) {
    if ($f.Name -match "이정우") {
        foreach ($m in $f.Items) {
            if ($m.Subject -match "SAP") {
                $s = $m.Subject
                $b = $m.Body.Substring(0,500)
                Write-Host "SUB: $s"
                Write-Host "BODY: $b"
                return
            }
        }
    }
}

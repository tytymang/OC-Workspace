
$ErrorActionPreference = "SilentlyContinue"
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6)
foreach ($f in $inbox.Folders) {
    if ($f.Name.Contains("이정우")) {
        foreach ($i in $f.Items) {
            if ($i.Subject.Contains("SAP")) {
                $i.Subject
                "---"
                $i.Body.Substring(0, 500)
                return
            }
        }
    }
}

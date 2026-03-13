
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $root = $namespace.Folders.Item(1)
    
    function FindFolder($parent, $name) {
        foreach ($f in $parent.Folders) {
            if ($f.Name -like "*$name*") { return $f }
            $sub = FindFolder $f $name
            if ($sub -ne $null) { return $sub }
        }
        return $null
    }

    $targetFolder = FindFolder $root "이정우"
    if ($targetFolder -eq $null) { throw "FOLDER_NOT_FOUND" }

    $targetMail = $null
    $items = $targetFolder.Items
    $items.Sort("[ReceivedTime]", $true)
    
    foreach ($item in $items) {
        if ($item.Subject -like "*SAP*" -and $item.Subject -like "*세미나*") {
            $targetMail = $item
            break
        }
    }

    if ($targetMail -ne $null) {
        $res = [PSCustomObject]@{
            Subject = $targetMail.Subject
            Body = $targetMail.Body
        }
        $res | ConvertTo-Json
    } else {
        "MAIL_NOT_FOUND"
    }
} catch {
    $_.Exception.Message
}

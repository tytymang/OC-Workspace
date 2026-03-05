
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $root = $namespace.Folders.Item(1)
    
    function Get-TargetFolder($parent) {
        foreach ($f in $parent.Folders) {
            if ($f.Name -like "*이정우*") { return $f }
            $sub = Get-TargetFolder $f
            if ($sub) { return $sub }
        }
        return $null
    }

    $folder = Get-TargetFolder $root
    if (-not $folder) { "FOLDER_NOT_FOUND"; exit }

    $items = $folder.Items
    $items.Sort("[ReceivedTime]", $true)
    foreach ($item in $items) {
        if ($item.Subject -like "*SAP*" -and $item.Subject -like "*세미나*") {
            $obj = [PSCustomObject]@{
                Subject = $item.Subject
                Body = $item.Body
            }
            $obj | ConvertTo-Json
            exit
        }
    }
    "MAIL_NOT_FOUND"
} catch {
    $_.Exception.Message
}

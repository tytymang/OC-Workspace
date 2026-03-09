
try {
    $outlook = New-Object -ComObject Outlook.Application
    $inbox = $outlook.GetNamespace("MAPI").GetDefaultFolder(6)
    $folder = $null
    foreach ($f in $inbox.Folders) {
        if ($f.Name -match "이정우") { $folder = $f; break }
    }
    
    if ($folder) {
        $found = $null
        foreach ($item in $folder.Items) {
            # SAP (83, 65, 80)
            if ($item.Subject -match "SAP") {
                $found = $item
                break
            }
        }
        if ($found) {
            $obj = [PSCustomObject]@{
                Subject = $found.Subject
                Body = $found.Body.Trim().Substring(0, 1000)
            }
            $obj | ConvertTo-Json
        } else { "MAIL_NOT_FOUND" }
    } else { "FOLDER_NOT_FOUND" }
} catch { $_.Exception.Message }

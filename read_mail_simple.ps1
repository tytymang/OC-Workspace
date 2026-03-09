
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $folder = $null
    foreach ($f in $inbox.Folders) {
        if ($f.Name -match "이정우") { $folder = $f; break }
    }
    
    if (-not $folder) { "FOLDER_NOT_FOUND"; exit }

    $items = $folder.Items
    $items.Sort("[ReceivedTime]", $true)
    $mails = @()
    $count = 0
    foreach ($item in $items) {
        if ($item.Subject -match "SAP" -and $count -lt 3) {
            $mails += [PSCustomObject]@{
                Subject = $item.Subject
                Body = $item.Body.Substring(0, [Math]::Min($item.Body.Length, 1500))
            }
            $count++
        }
    }
    $mails | ConvertTo-Json
} catch { $_.Exception.Message }

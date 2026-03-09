
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $targetFolder = $null
    foreach ($f in $inbox.Folders) {
        if ($f.Name -match "이정우") {
            $targetFolder = $f
            break
        }
    }
    
    if (-not $targetFolder) { Write-Output "FOLDER_NOT_FOUND"; exit }

    $items = $targetFolder.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $foundMail = $null
    foreach ($item in $items) {
        # Search for "SAP" and "Connect" in subject, looking for Feb 10th (approx)
        if ($item.Subject -match "SAP" -and $item.Subject -match "Connect") {
            # Check if received around Feb 10 (02-10)
            if ($item.ReceivedTime.ToString("MM-dd") -eq "02-10") {
                $foundMail = $item
                break
            }
        }
    }

    # If not found on exact date, take the latest matching subject
    if (-not $foundMail) {
        foreach ($item in $items) {
            if ($item.Subject -match "SAP" -and $item.Subject -match "Connect") {
                $foundMail = $item
                break
            }
        }
    }

    if ($foundMail) {
        $res = [PSCustomObject]@{
            Subject = $foundMail.Subject
            Received = $foundMail.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            Body = $foundMail.Body.Trim().Substring(0, [Math]::Min($foundMail.Body.Length, 1000))
        }
        $res | ConvertTo-Json
    } else {
        Write-Output "MAIL_NOT_FOUND"
    }
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}

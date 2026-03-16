
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $deletedItems = $namespace.GetDefaultFolder(3) # olFolderDeletedItems
    $items = $deletedItems.Items
    $items.Sort("[ReceivedTime]", $true)

    $found = $items | Where-Object { $_.Subject -match "Trip.com" -or $_.Body -match "Trip.com" } | Select-Object -First 5

    if ($null -eq $found) {
        Write-Host "NOT_FOUND_IN_DELETED"
    } else {
        foreach ($item in $found) {
            Write-Host "---"
            Write-Host "Subject: $($item.Subject)"
            Write-Host "Received: $($item.ReceivedTime)"
            $body = $item.Body
            Write-Host "Body: $($body.Substring(0, [Math]::Min(2000, $body.Length)))"
        }
    }
} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($null -ne $outlook) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
}

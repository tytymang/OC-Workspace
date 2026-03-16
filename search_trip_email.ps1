
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6) # olFolderInbox
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)

    # Search for Trip.com related emails
    $found = $items | Where-Object { $_.Subject -match "Trip.com" -or $_.Body -match "Trip.com" } | Select-Object -First 5

    if ($null -eq $found) {
        Write-Host "NOT_FOUND"
    } else {
        foreach ($item in $found) {
            Write-Host "---"
            Write-Host "Subject: $($item.Subject)"
            Write-Host "Received: $($item.ReceivedTime)"
            # Extract basic flight info if possible using regex
            $body = $item.Body
            Write-Host "Body: $($body.Substring(0, [Math]::Min(1000, $body.Length)))"
        }
    }
} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($null -ne $outlook) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
}

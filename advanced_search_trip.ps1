
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $search = $outlook.AdvancedSearch("'\\$($namespace.DefaultStore.DisplayName)'", "urn:schemas:httpmail:subject LIKE '%Trip.com%' OR urn:schemas:httpmail:textdescription LIKE '%Trip.com%'", $true)
    
    # Wait for search to complete (simple poll)
    Start-Sleep -Seconds 5
    
    if ($search.Results.Count -eq 0) {
        Write-Host "NOT_FOUND"
    } else {
        foreach ($item in $search.Results) {
            Write-Host "---"
            Write-Host "Subject: $($item.Subject)"
            Write-Host "Received: $($item.ReceivedTime)"
            Write-Host "Body: $($item.Body.Substring(0, [Math]::Min(1000, $item.Body.Length)))"
        }
    }
} catch {
    Write-Error $_.Exception.Message
}

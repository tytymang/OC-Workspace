
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace("MAPI")
    $inbox = $ns.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $targetSubject = "Error ZSCMR3000"
    $found = $false
    
    for ($i = 1; $i -le 20; $i++) {
        $item = $items.Item($i)
        if ($item.Subject -like "*$targetSubject*") {
            Write-Output "--- Subject: $($item.Subject) ---"
            Write-Output "--- From: $($item.SenderName) ---"
            Write-Output "--- Body Start ---"
            Write-Output $item.Body.Substring(0, [Math]::Min(2000, $item.Body.Length))
            Write-Output "--- Body End ---"
            $found = $true
            break
        }
    }
} catch {
    Write-Output $_.Exception.Message
}


try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $mail = $inbox.Items | Select-Object -First 1
    Write-Output "LATEST_MAIL: $($mail.ReceivedTime) - $($mail.Subject)"
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}

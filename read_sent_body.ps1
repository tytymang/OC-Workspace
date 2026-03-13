
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $sentFolder = $namespace.GetDefaultFolder(5)
    
    $latestMail = $sentFolder.Items | Sort-Object SentOn -Descending | Select-Object -First 1
    
    if ($latestMail -ne $null) {
        $result = [PSCustomObject]@{
            Subject = $latestMail.Subject
            Body = $latestMail.Body
        }
        $result | ConvertTo-Json
    } else {
        "MAIL_NOT_FOUND"
    }
} catch {
    Write-Error $_.Exception.Message
}

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)

# Search for recent BPM notification emails that might contain the URL
$bpmMails = $inbox.Items.Restrict("[SenderName] = 'SSCBPM'")
$bpmMails.Sort("[ReceivedTime]", $true)

$results = @()
foreach ($mail in $bpmMails | Select-Object -First 5) {
    $results += @{
        Subject = $mail.Subject
        BodySnippet = if ($mail.Body.Length -gt 500) { $mail.Body.Substring(0, 500) } else { $mail.Body }
    }
}

$results | ConvertTo-Json | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\bpm_mail_sample.json" -Encoding UTF8

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)

    $results = @()
    # Broaden search to include typical BPM subjects if SenderName isn't an exact match
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)

    $count = 0
    foreach ($mail in $items) {
        if ($mail.SenderName -match "BPM" -or $mail.Subject -match "BPM") {
            $results += @{
                Subject = $mail.Subject
                Body = if ($mail.Body -and $mail.Body.Length -gt 1000) { $mail.Body.Substring(0, 1000) } else { $mail.Body }
            }
            $count++
        }
        if ($count -ge 5) { break }
    }
    $results | ConvertTo-Json | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\bpm_mail_sample_v2.json" -Encoding UTF8
} catch {
    $_.Exception.Message | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\bpm_extract_err.txt"
}

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)

    $results = @()
    $count = 0
    foreach ($mail in $items) {
        if ($mail.Subject -match "BPM" -and $mail.Body -match "http") {
            $results += @{
                Subject = $mail.Subject
                Body = $mail.Body
            }
            $count++
        }
        if ($count -ge 10) { break }
    }
    $results | ConvertTo-Json | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\bpm_links_search.json" -Encoding UTF8
} catch {
    $_.Exception.Message | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\bpm_links_err.txt"
}

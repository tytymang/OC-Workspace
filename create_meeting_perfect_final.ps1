
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $subject = [char]50 + [char]50900 + [char]32 + [char]47560 + [char]44048 + [char]32 + [char]54924 + [char]51032 + [char]32 + [char]54869 + [char]51064
    $body = [char]50 + [char]50900 + [char]32 + [char]47560 + [char]44048 + [char]32 + [char]50756 + [char]47308 + [char]32 + [char]49324 + [char]54637 + [char]32 + [char]54869 + [char]51064 + [char]51012 + [char]32 + [char]50948 + [char]54620 + [char]32 + [char]54924 + [char]51032 + [char]32 + [char]51077 + [char]45768 + [char]45796 + [char]46
    $appt = $outlook.CreateItem(1)
    $appt.Subject = $subject
    $appt.Body = $body
    $appt.Location = "https://sskv.webex.com/sskv/j.php?MTID=mb9e45aa44dbb35ccfc41197aabb63e1b"
    $appt.Start = Get-Date -Year 2026 -Month 3 -Day 5 -Hour 10 -Minute 30 -Second 0
    $appt.End = Get-Date -Year 2026 -Month 3 -Day 5 -Hour 12 -Minute 0 -Second 0
    $appt.MeetingStatus = 1
    $recipients = @("huyentrang@seoulviosys.com", "703976@seoulsemicon.com", "wen1228@seoulsemicon.com")
    foreach ($addr in $recipients) {
        $recip = $appt.Recipients.Add($addr)
        $recip.Type = 1
    }
    $appt.Recipients.ResolveAll()
    $appt.Display()
    "SUCCESS"
} catch {
    $_.Exception.Message | Out-File -FilePath "C:\Users\307984\.openclaw\workspace\final_fix_v3.log"
    throw $_
}

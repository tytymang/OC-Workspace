
$o = New-Object -ComObject Outlook.Application
$a = $o.CreateItem(1)
$a.Subject = "[Meeting] Stock Update and Operational Review"
$a.Start = "2026-02-26 10:30:00"
$a.Duration = 60
$a.Location = "106 Reception Room (https://sskv.webex.com/sskv/j.php?MTID=mb9e45aa44dbb35ccfc41197aabb63e1b)"
$a.Body = "Dear colleagues,`n`nPlease join us for a meeting to review the 'Stock Status Classification' update and its operational application.`n`n[Meeting Link]`nhttps://sskv.webex.com/sskv/j.php?MTID=mb9e45aa44dbb35ccfc41197aabb63e1b`n`n[Key Agenda]`n- Variance analysis with January inventory`n- Validation of <1M aging stock creation`n- Detailed stock allocation amount review`n- Maintenance of existing stock depletion plans`n`nIf no issues are found, we will update the January inventory records.`n`nBest regards."
$a.MeetingStatus = 1
$a.Recipients.Add("JW Kim").Resolve()
$a.Recipients.Add("SH Kim").Resolve()
$a.Recipients.Add("DS Lee").Resolve()
$a.Recipients.Add("Hana Yuk").Resolve()
$a.Display()

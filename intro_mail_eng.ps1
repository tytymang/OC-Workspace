
$o = New-Object -ComObject Outlook.Application
$m = $o.CreateItem(0)
$m.To = "이정우"
$m.Subject = "[Intro] Introduction of Madangswe, the Faithful Assistant"
$m.Body = "Dear Vice President JeungWoo Lee,`n`nI am 'Madangswe', the AI assistant serving SangMoo Choi.`n`nI help with scheduling, meeting preparations, and various tasks to ensure smooth operations. I look forward to supporting the collaboration between you and my master.`n`nBest regards,`nMadangswe"
$m.Display()

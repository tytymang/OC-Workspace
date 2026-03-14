$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6)
$sent = $ns.GetDefaultFolder(5)

$lastWeek = (Get-Date).AddDays(-7)
$filter = "[ReceivedTime] >= '" + $lastWeek.ToString("g") + "'"

$receivedMails = $inbox.Items.Restrict($filter)
$sentMails = $sent.Items.Restrict("[SentOn] >= '" + $lastWeek.ToString("g") + "'")

$sentSubjects = @{}
foreach ($s in $sentMails) {
    $sub = $s.Subject -replace "^RE:\s*", "" -replace "^FW:\s*", ""
    $sentSubjects[$sub.Trim()] = $s.SentOn
}

$results = @()
foreach ($m in $receivedMails) {
    $cleanSub = $m.Subject -replace "^RE:\s*", "" -replace "^FW:\s*", ""
    $cleanSub = $cleanSub.Trim()
    
    # Check if replied
    if (-not $sentSubjects.ContainsKey($cleanSub)) {
        # Heuristic for "needs reply": contains ?, 부탁, 회신, 확인, 요청, 등
        if ($m.Body -match "\?|부탁|회신|확인|요청|검토|의견|컨펌") {
             $results += [PSCustomObject]@{
                Received = $m.ReceivedTime.ToString("MM-dd HH:mm")
                Sender = $m.SenderName
                Subject = $m.Subject
                Preview = if ($m.Body.Length -gt 50) { $m.Body.Substring(0, 50).Replace("`r`n", " ") } else { $m.Body.Replace("`r`n", " ") }
             }
        }
    }
}

$results | ConvertTo-Json

$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6)
$sentFolder = $ns.GetDefaultFolder(5)

$lastWeek = (Get-Date).AddDays(-7)
$receivedMails = $inbox.Items | Where-Object { $_.ReceivedTime -ge $lastWeek }
$sentMails = $sentFolder.Items | Where-Object { $_.SentOn -ge $lastWeek }

$sentSubjects = @{}
foreach ($s in $sentMails) {
    $sub = $s.Subject.Replace("RE: ", "").Replace("FW: ", "").Trim()
    $sentSubjects[$sub] = $s.SentOn
}

$results = @()
foreach ($m in $receivedMails) {
    $sub = $m.Subject.Replace("RE: ", "").Replace("FW: ", "").Trim()
    
    if (-not $sentSubjects.ContainsKey($sub)) {
        $body = $m.Body
        # Use simpler matching for non-ASCII
        if ($body.Contains("?") -or $body.Contains("부탁") -or $body.Contains("회신") -or $body.Contains("확인") -or $body.Contains("요청") -or $body.Contains("검토")) {
             $obj = New-Object PSObject
             $obj | Add-Member NoteProperty Received ($m.ReceivedTime.ToString("MM-dd HH:mm"))
             $obj | Add-Member NoteProperty Sender ($m.SenderName)
             $obj | Add-Member NoteProperty Subject ($m.Subject)
             $results += $obj
        }
    }
}

$results | ConvertTo-Json

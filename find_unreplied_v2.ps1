$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6)
$sentFolder = $ns.GetDefaultFolder(5)

$lastWeek = (Get-Date).AddDays(-7)
$receivedMails = $inbox.Items | ? { $_.ReceivedTime -ge $lastWeek }
$sentMails = $sentFolder.Items | ? { $_.SentOn -ge $lastWeek }

$sentSubjects = @{}
foreach ($s in $sentMails) {
    $sub = $s.Subject.Replace("RE: ", "").Replace("FW: ", "").Trim()
    $sentSubjects[$sub] = $s.SentOn
}

$results = @()
foreach ($m in $receivedMails) {
    $cleanSub = $m.Subject.Replace("RE: ", "").Replace("FW: ", "").Trim()
    
    if (-not $sentSubjects.ContainsKey($cleanSub)) {
        $body = $m.Body
        if ($body -match "\?|부탁|회신|확인|요청|검토|의견|컨펌") {
             $obj = New-Object PSObject
             $obj | Add-Member NoteProperty Received ($m.ReceivedTime.ToString("yyyy-MM-dd HH:mm"))
             $obj | Add-Member NoteProperty Sender ($m.SenderName)
             $obj | Add-Member NoteProperty Subject ($m.Subject)
             $results += $obj
        }
    }
}

$results | ConvertTo-Json

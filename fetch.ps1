$o = New-Object -ComObject Outlook.Application
$n = $o.GetNamespace("MAPI")
$f = $n.GetDefaultFolder(6).Folders.Item("중요 업무")
$s = @("김하영", "이수정")
$r = @()
foreach ($sn in $s) {
    $m = $f.Items | Where-Object { $_.SenderName -like "*$sn*" -and $_.Attachments.Count -gt 0 } | Sort-Object ReceivedTime -Descending | Select-Object -First 1
    if ($m) {
        foreach ($a in $m.Attachments) {
            $p = Join-Path $env:TEMP $a.FileName
            $a.SaveAsFile($p)
            $r += [PSCustomObject]@{Sender=$sn;File=$a.FileName;Path=$p}
        }
    }
}
$j = $r | ConvertTo-Json
[System.IO.File]::WriteAllText("C:\Users\307984\.openclaw\workspace\mail_attachments.json", $j, [System.Text.Encoding]::Unicode)

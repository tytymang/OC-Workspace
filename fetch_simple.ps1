$o = New-Object -ComObject Outlook.Application
$n = $o.GetNamespace('MAPI')
$f = $n.GetDefaultFolder(6).Folders.Item('중요 업무')
$s = @('김하영', '이수정')
$r = @()
foreach ($sn in $s) {
    $m = $f.Items | Where-Object { $_.SenderName -like "*$sn*" -and $_.Attachments.Count -gt 0 } | Sort-Object ReceivedTime -Descending | Select-Object -First 1
    if ($m) {
        foreach ($a in $m.Attachments) {
            $p = Join-Path $env:TEMP $a.FileName
            $a.SaveAsFile($p)
            $obj = New-Object PSObject
            $obj | Add-Member NoteProperty Sender $sn
            $obj | Add-Member NoteProperty File $a.FileName
            $obj | Add-Member NoteProperty Path $p
            $r += $obj
        }
    }
}
$j = $r | ConvertTo-Json
$j > out.json

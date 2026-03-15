
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)
$items = $inbox.Items
$items.Sort("[ReceivedTime]", $true)

$res = ""
for ($i=1; $i -le 5; $i++) {
    $m = $items.Item($i)
    $res += "[$($m.ReceivedTime.ToString('MM/dd HH:mm'))] $($m.SenderName): $($m.Subject)`r`n"
}

[System.IO.File]::WriteAllText("C:\Users\307984\.openclaw\workspace\mails_utf16.txt", $res, [System.Text.Encoding]::Unicode)

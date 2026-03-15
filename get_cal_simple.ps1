
$outlook = New-Object -ComObject Outlook.Application
$calendar = $outlook.GetNamespace("MAPI").GetDefaultFolder(9)
$today = (Get-Date).Date
$tomorrow = $today.AddDays(1)

$res = "--- 오늘 일정 ---`r`n"
foreach ($item in $calendar.Items) {
    if ($item.Start -ge $today -and $item.Start -lt $tomorrow) {
        $res += "[$($item.Start.ToString('HH:mm'))] $($item.Subject)`r`n"
    }
}
[System.IO.File]::WriteAllText("C:\Users\307984\.openclaw\workspace\cal_utf16.txt", $res, [System.Text.Encoding]::Unicode)

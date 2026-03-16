
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $res = ""
    $count = 0
    foreach ($m in $items) {
        if ($m.Subject -match "Trip.com" -or $m.SenderName -match "Trip.com" -or $m.Body -match "Trip.com") {
            $res += "[$($m.ReceivedTime.ToString('yyyy-MM-dd HH:mm'))] $($m.SenderName): $($m.Subject)`r`n"
            $res += "BODY_START`r`n$($m.Body)`r`nBODY_END`r`n"
        }
    }
    
    if ($res -eq "") { $res = "Trip.com 관련 메일을 찾을 수 없습니다." }
    [System.IO.File]::WriteAllText("C:\Users\307984\.openclaw\workspace\trip_mails.txt", $res, [System.Text.Encoding]::Unicode)
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}

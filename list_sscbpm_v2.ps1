
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $res = ""
    foreach ($m in $items) {
        if ($m.SenderName -like "*SSCBPM*" -or $m.Subject -like "*SSCBPM*") {
            $res += "[$($m.ReceivedTime.ToString('yyyy-MM-dd HH:mm'))] $($m.SenderName): $($m.Subject)`r`n"
        }
    }
    
    if ($res -eq "") { $res = "검색된 메일이 없습니다." }
    
    [System.IO.File]::WriteAllText("C:\Users\307984\.openclaw\workspace\sscbpm_list.txt", $res, [System.Text.Encoding]::Unicode)
} catch {
    [System.IO.File]::WriteAllText("C:\Users\307984\.openclaw\workspace\sscbpm_list.txt", "ERROR: $($_.Exception.Message)", [System.Text.Encoding]::Unicode)
}

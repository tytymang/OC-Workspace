
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # "3/18" 이나 "4/3" 이 포함된 메일 검색
    $res = ""
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    foreach ($m in $items) {
        if ($m.ReceivedTime -lt (Get-Date).AddDays(-14)) { break }
        if ($m.Subject -like "*3/18*" -or $m.Subject -like "*4/3*" -or $m.Body -like "*3/18*" -or $m.Body -like "*4/3*") {
             $res += "[$($m.ReceivedTime.ToString('yyyy-MM-dd HH:mm'))] $($m.SenderName): $($m.Subject)`r`n"
        }
    }
    
    [System.IO.File]::WriteAllText("C:\Users\307984\.openclaw\workspace\date_search.txt", $res, [System.Text.Encoding]::Unicode)
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}

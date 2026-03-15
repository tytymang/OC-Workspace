
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # SSCBPM이 포함된 발신자 메일 검색
    $filter = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0042001E"" LIKE '%SSCBPM%'"
    $items = $inbox.Items.Restrict($filter)
    $items.Sort("[ReceivedTime]", $true)
    
    $res = ""
    foreach ($m in $items) {
        $res += "[$($m.ReceivedTime.ToString('yyyy-MM-dd HH:mm'))] $($m.SenderName): $($m.Subject)`r`n"
    }
    
    if ($res -eq "") { $res = "검색된 메일이 없습니다." }
    
    [System.IO.File]::WriteAllText("C:\Users\307984\.openclaw\workspace\sscbpm_list.txt", $res, [System.Text.Encoding]::Unicode)
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}

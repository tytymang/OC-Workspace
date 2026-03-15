
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # 지운 편지함(olFolderDeletedItems = 3)도 확인
    $deleted = $namespace.GetDefaultFolder(3)
    
    $res = "--- [받은 편지함] SSCBPM 메일 ---`r`n"
    foreach ($m in $inbox.Items) {
        if ($m.SenderName -like "*SSCBPM*" -or $m.Subject -like "*SSCBPM*") {
            $res += "[INBOX] [$($m.ReceivedTime.ToString('yyyy-MM-dd HH:mm'))] $($m.SenderName): $($m.Subject)`r`n"
        }
    }
    
    $res += "`r`n--- [지운 편지함] SSCBPM 메일 ---`r`n"
    foreach ($m in $deleted.Items) {
        if ($m.SenderName -like "*SSCBPM*" -or $m.Subject -like "*SSCBPM*") {
            $res += "[DELETED] [$($m.ReceivedTime.ToString('yyyy-MM-dd HH:mm'))] $($m.SenderName): $($m.Subject)`r`n"
        }
    }
    
    [System.IO.File]::WriteAllText("C:\Users\307984\.openclaw\workspace\sscbpm_full_list.txt", $res, [System.Text.Encoding]::Unicode)
} catch {
    [System.IO.File]::WriteAllText("C:\Users\307984\.openclaw\workspace\sscbpm_full_list.txt", "ERROR: $($_.Exception.Message)", [System.Text.Encoding]::Unicode)
}

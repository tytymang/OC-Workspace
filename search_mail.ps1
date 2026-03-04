
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

try {
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace('MAPI')
    $Inbox = $Namespace.GetDefaultFolder(6)
    $Items = $Inbox.Items
    
    # 최근 메일 순으로 정렬
    $Items.Sort("[ReceivedTime]", $true)
    
    $Results = @()
    $filter = "[Subject] >= '2026' AND [Subject] <= '2027'"
    $RestrictedItems = $Items.Restrict($filter)
    $RestrictedItems.Sort("[ReceivedTime]", $true)

    foreach ($Item in $RestrictedItems) {
        if ($Item.Subject -like "*2분기*" -and $Item.Subject -like "*토요*") {
            $Results += [PSCustomObject]@{
                Subject = $Item.Subject
                Body    = $Item.Body
            }
            break
        }
    }
    
    if ($Results.Count -eq 0) {
        Write-Host "No matching emails found."
    } else {
        $Results | ConvertTo-Json
    }
} catch {
    Write-Error $_.Exception.Message
} finally {
    if ($Outlook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null }
}

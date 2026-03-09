
$ErrorActionPreference = "Stop"
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $sentFolder = $namespace.GetDefaultFolder(5)
    $items = $sentFolder.Items
    
    # 2월 마감 완료 확인 제목의 최신 보낸 메일 검색
    $targetMail = $null
    # Get top 50 sent items to find the match quickly
    $count = [Math]::Min($items.Count, 50)
    for ($i = $items.Count; $i -gt ($items.Count - $count); $i--) {
        $item = $items.Item($i)
        if ($item.Subject -like "*2월 마감 완료 확인*") {
            $targetMail = $item
            break
        }
    }
    
    if ($targetMail -ne $null) {
        $recipients = @()
        foreach ($recipient in $targetMail.Recipients) {
            if ($recipient.Type -eq 1) { # olTo
                $recipients += $recipient.Name
            }
        }
        $recipients | ConvertTo-Json
    } else {
        "MAIL_NOT_FOUND"
    }
} catch {
    $_.Exception.Message | Out-File -FilePath "error_mail.log"
    throw $_
}

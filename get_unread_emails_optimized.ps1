$ErrorActionPreference = "Stop"
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$results = @()

function Get-UnreadEmails($folder) {
    # Skip Deleted Items or Junk if needed, but instruction said "ALL folders"
    try {
        $unreadItems = $folder.Items.Restrict("[UnRead] = true")
        foreach ($item in $unreadItems) {
            if ($null -ne $item) {
                $summary = ""
                # Attempt to get body; handle non-mail items like meeting requests
                try {
                    if ($null -ne $item.Body) {
                        $summary = if ($item.Body.Length -gt 100) { $item.Body.Substring(0, 100).Replace("`r`n", " ").Replace("`n", " ") + "..." } else { $item.Body.Replace("`r`n", " ").Replace("`n", " ") }
                    }
                } catch {
                    $summary = "(본문을 가져올 수 없습니다: " + $item.MessageClass + ")"
                }
                
                $obj = [PSCustomObject]@{
                    Folder   = $folder.Name
                    Sender   = $item.SenderName
                    SentTime = $item.SentOn.ToString("yyyy-MM-dd HH:mm")
                    Subject  = $item.Subject
                    Summary  = $summary
                }
                $results += $obj
            }
        }
    } catch {
        # Some folders might not have Items or items might be of different type
    }

    foreach ($subFolder in $folder.Folders) {
        Get-UnreadEmails($subFolder)
    }
}

foreach ($root in $namespace.Folders) {
    Get-UnreadEmails($root)
}

$results | ConvertTo-Json -Compress

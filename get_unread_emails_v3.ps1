$ErrorActionPreference = "Continue"
try {
    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $namespace.Logon()
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    function Get-UnreadEmails($folder) {
        # Skip certain folder types if needed
        # Default folders are: 6 (Inbox), 3 (Deleted), 23 (Junk) etc.
        try {
            $unreadItems = $folder.Items.Restrict("[UnRead] = true")
            foreach ($item in $unreadItems) {
                if ($null -ne $item -and ($item.MessageClass -like "IPM.Note*")) {
                    $sender = try { $item.SenderName } catch { "(N/A)" }
                    $subject = try { $item.Subject } catch { "(N/A)" }
                    $sentTime = try { $item.SentOn.ToString("yyyy-MM-dd HH:mm") } catch { "(N/A)" }
                    $body = try { $item.Body } catch { "" }
                    $summary = ""
                    if ($null -ne $body) {
                        $summary = $body.Substring(0, [Math]::Min(100, $body.Length)).Replace("`r`n", " ").Replace("`n", " ")
                        if ($body.Length -gt 100) { $summary += "..." }
                    }
                    
                    $results.Add([PSCustomObject]@{
                        Folder   = $folder.Name
                        Sender   = $sender
                        SentTime = $sentTime
                        Subject  = $subject
                        Summary  = $summary
                    })
                }
            }
        } catch {
            # Silent fail for specific properties
        }

        foreach ($subFolder in $folder.Folders) {
            Get-UnreadEmails($subFolder)
        }
    }

    foreach ($root in $namespace.Folders) {
        Get-UnreadEmails($root)
    }

    if ($results.Count -gt 0) {
        $results | ConvertTo-Json -Compress
    } else {
        "[]"
    }
} catch {
    Write-Error "CRITICAL: $($_.Exception.Message)"
}

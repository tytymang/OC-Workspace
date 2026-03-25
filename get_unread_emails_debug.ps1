$ErrorActionPreference = "Continue"
try {
    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $namespace.Logon()
    $results = @()

    function Get-UnreadEmails($folder) {
        Write-Host "Scanning: $($folder.FolderPath)"
        try {
            # Try to get unread items
            $unreadItems = $folder.Items.Restrict("[UnRead] = true")
            foreach ($item in $unreadItems) {
                if ($null -ne $item) {
                    $sender = try { $item.SenderName } catch { "(N/A)" }
                    $subject = try { $item.Subject } catch { "(N/A)" }
                    $sentTime = try { $item.SentOn.ToString("yyyy-MM-dd HH:mm") } catch { "(N/A)" }
                    $summary = try { 
                        if ($null -ne $item.Body) {
                            $item.Body.Substring(0, [Math]::Min(100, $item.Body.Length)).Replace("`r`n", " ").Replace("`n", " ") + "..."
                        } else { "" }
                    } catch { "(본문을 가져올 수 없습니다)" }
                    
                    $results += [PSCustomObject]@{
                        Folder   = $folder.Name
                        Sender   = $sender
                        SentTime = $sentTime
                        Subject  = $subject
                        Summary  = $summary
                    }
                }
            }
        } catch {
            Write-Warning "Failed to read folder $($folder.FolderPath): $($_.Exception.Message)"
        }

        foreach ($subFolder in $folder.Folders) {
            Get-UnreadEmails($subFolder)
        }
    }

    foreach ($root in $namespace.Folders) {
        Get-UnreadEmails($root)
    }

    $results | ConvertTo-Json -Compress
} catch {
    Write-Error "CRITICAL: $($_.Exception.Message)"
}

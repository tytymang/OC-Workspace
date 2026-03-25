$ErrorActionPreference = "Stop"
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$results = @()

function Get-UnreadEmails($folder) {
    # Skip Deleted Items or Junk if needed, but the instruction said "ALL folders"
    try {
        foreach ($item in $folder.Items) {
            if ($null -ne $item -and $item.UnRead -eq $true) {
                $summary = ""
                if ($null -ne $item.Body) {
                    $summary = if ($item.Body.Length -gt 100) { $item.Body.Substring(0, 100).Replace("`r`n", " ").Replace("`n", " ") + "..." } else { $item.Body.Replace("`r`n", " ").Replace("`n", " ") }
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
        # Some folders might not have Items or items might be of different type (e.g. Calendar items in a calendar folder)
    }

    foreach ($subFolder in $folder.Folders) {
        Get-UnreadEmails($subFolder)
    }
}

foreach ($root in $namespace.Folders) {
    Get-UnreadEmails($root)
}

$results | ConvertTo-Json -Compress

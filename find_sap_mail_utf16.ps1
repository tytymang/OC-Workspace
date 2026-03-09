
$ErrorActionPreference = "Stop"

# Force output encoding to UTF-8 for console
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # Target Folder: 이정우 부사장 (Partial Match)
    $targetFolder = $null
    foreach ($f in $inbox.Folders) {
        if ($f.Name -match "이정우") {
            $targetFolder = $f
            break
        }
    }
    
    if (-not $targetFolder) {
        Write-Host "FOLDER_NOT_FOUND"
        exit
    }

    # Search for SAP mail
    $foundMail = $null
    foreach ($item in $targetFolder.Items) {
        if ($item.Subject -match "SAP") {
            $foundMail = $item
            break
        }
    }

    if ($foundMail) {
        $subject = $foundMail.Subject
        $body = $foundMail.Body
        if ($body.Length -gt 1000) { $body = $body.Substring(0, 1000) }
        
        # Output JSON manually to avoid encoding issues with ConvertTo-Json
        Write-Host "{""Subject"":""$subject"",""Body"":""$body""}"
    } else {
        Write-Host "MAIL_NOT_FOUND"
    }
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
}

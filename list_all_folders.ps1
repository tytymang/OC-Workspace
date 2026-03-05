
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    function Get-Folders($parent, $indent) {
        foreach ($f in $parent.Folders) {
            Write-Host "$indent$($f.Name)"
            Get-Folders $f "$indent  "
        }
    }

    $root = $namespace.Folders.Item(1)
    Get-Folders $root ""
} catch { $_.Exception.Message }

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    $secretaryFolder = $inbox.Folders.Item(7)
    $kimFolder = $secretaryFolder.Folders.Item(2)
    
    $items = $kimFolder.Items
    $results = @()
    for ($i = 1; $i -le 5; $i++) {
        $item = $items.Item($i)
        $obj = New-Object PSObject
        $obj | Add-Member -MemberType NoteProperty -Name "Index" -Value $i
        $obj | Add-Member -MemberType NoteProperty -Name "Subject" -Value $item.Subject
        $results += $obj
    }
    $results | ConvertTo-Json
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
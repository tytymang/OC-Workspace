$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")
$GAL = $Namespace.GetGlobalAddressList()
$Entries = $GAL.AddressEntries

# Try to find exactly "육하나"
try {
    $Person = $Entries.Item("육하나")
    if ($Person -ne $null) {
        $ExUser = $Person.GetExchangeUser()
        $alias = $ExUser.Alias
        $jobTitle = $ExUser.JobTitle
        $dept = $ExUser.Department
        Write-Output "Found: $alias | $jobTitle | $dept"
    } else {
        Write-Output "Person '육하나' not found directly."
    }
} catch {
    Write-Output "Error searching: $_"
}

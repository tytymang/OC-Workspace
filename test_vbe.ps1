$ErrorActionPreference = 'Stop'
$Outlook = New-Object -ComObject Outlook.Application
try {
    $vbe = $Outlook.LanguageSettings
    Write-Output "LanguageSettings works."
    $vbe = $Outlook.VBE
    if ($vbe -eq $null) { Write-Output "VBE is null" } else { Write-Output "VBE exists" }
} catch {
    Write-Output "Error: $_"
}

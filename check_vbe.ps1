$ErrorActionPreference = 'Stop'
$Outlook = New-Object -ComObject Outlook.Application
Write-Output "Version: $($Outlook.Version)"
try {
    $vbe = $Outlook.VBE
    Write-Output "VBE object available."
    $proj = $vbe.ActiveVBProject
    Write-Output "ActiveVBProject accessible."
} catch {
    Write-Output "VBE Access Denied: $_"
}

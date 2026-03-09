$ErrorActionPreference = 'Stop'
$Outlook = New-Object -ComObject Outlook.Application
$vbe = $Outlook.VBE
$proj = $vbe.ActiveVBProject

foreach ($comp in $proj.VBComponents) {
    Write-Output "Component Name: $($comp.Name), Type: $($comp.Type)"
}

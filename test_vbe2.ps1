$app = New-Object -ComObject Outlook.Application
try {
    $vbe = $app.VBE
    Write-Output "VBE: $vbe"
    $proj = $vbe.ActiveVBProject
    Write-Output "Project Name: $($proj.Name)"
} catch {
    Write-Output "Error: $_"
}

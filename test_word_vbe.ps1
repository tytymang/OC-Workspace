$ErrorActionPreference = 'Stop'
$w = New-Object -ComObject Word.Application
try {
    $vbe = $w.VBE
    Write-Output "Word VBE Access OK"
} catch {
    Write-Output "Word VBE Access Denied: $_"
}
$w.Quit()

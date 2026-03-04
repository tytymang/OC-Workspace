
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Outlook = New-Object -ComObject Outlook.Application
$Items = $Outlook.GetNamespace('MAPI').GetDefaultFolder(6).Items
$Items.Sort("[ReceivedTime]", $true)

$i = 0
foreach ($Item in $Items) {
    if ($Item.Subject -like "*2026*") {
        if ($i -eq 8) {
            $folder = Join-Path (Get-Location) "mail_images"
            if (!(Test-Path $folder)) { New-Item -ItemType Directory -Path $folder }
            
            $attachCount = 0
            foreach ($at in $Item.Attachments) {
                $safeName = $at.FileName -replace '[^a-zA-Z0-9\._-]', '_'
                $path = Join-Path $folder "image_$($attachCount)_$($safeName)"
                $at.SaveAsFile($path)
                Write-Host "SAVED: $path"
                $attachCount++
            }
            break
        }
        $i++
    }
}

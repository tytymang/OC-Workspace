
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)
$folder = $inbox.Folders | Where-Object { $_.Name -match "이정우" }
if ($folder) {
    $mail = $folder.Items | Sort-Object ReceivedTime -Descending | Select-Object -First 1
    if ($mail) {
        $mail.Subject
        $mail.Body
    }
}

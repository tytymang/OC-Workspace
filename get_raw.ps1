
$ErrorActionPreference = "SilentlyContinue"
$outlook = New-Object -ComObject Outlook.Application
$inbox = $outlook.GetNamespace("MAPI").GetDefaultFolder(6)
$folder = $null
foreach ($f in $inbox.Folders) {
    if ($f.Name -match "이정우") { $folder = $f; break }
}
if ($folder) {
    foreach ($item in $folder.Items) {
        if ($item.Subject -match "SAP") {
            "SUBJECT: " + $item.Subject
            "---BODY---"
            $item.Body.Trim().Substring(0, 1000)
            break
        }
    }
} else { "FOLDER_NOT_FOUND" }

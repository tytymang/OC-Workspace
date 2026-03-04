$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$folder = $namespace.GetDefaultFolder(9) # olFolderCalendar
$items = $folder.Items
$target = $items | Where-Object { $_.Subject -like "SCM *" }
foreach ($item in $target) {
    $item.Delete()
}
Write-Output "Broken appointments deleted."

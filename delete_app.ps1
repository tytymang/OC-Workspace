$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$folder = $namespace.GetDefaultFolder(9)
$items = $folder.Items
$target = $items | Where-Object { $_.Subject -eq "SCM 구축 협의" }
foreach ($item in $target) {
    $item.Delete()
}
Write-Output "Appointment deleted."

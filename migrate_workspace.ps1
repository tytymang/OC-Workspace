
$src = "C:\Users\307984\.openclaw\workspace"
$dest = "C:\Users\307984\Documents\openclaw"
if (!(Test-Path $dest)) { 
    New-Item -ItemType Directory -Path $dest -Force 
}
Copy-Item -Path "$src\*" -Destination $dest -Recurse -Force
Write-Output "Workspace files copied to $dest"

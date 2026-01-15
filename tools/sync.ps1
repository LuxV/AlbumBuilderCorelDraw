param([string]$Project)

$root = Resolve-Path "$PSScriptRoot\.."
$p = "$root\projects\$Project"

Remove-Item "$p\build" -Recurse -Force -ErrorAction SilentlyContinue
New-Item "$p\build" -ItemType Directory | Out-Null

Copy-Item "$p\source\*.bas" "$p\build" -Force
Copy-Item "$p\forms\*.frm" "$p\build" -Force
Copy-Item "$p\forms\*.frx" "$p\build" -Force


Write-Host "Build ready for import into Corel."
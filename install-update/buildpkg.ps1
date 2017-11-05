Write-Host "Ensure that the project has already been compiled under relase mode, this packager relies on the pre-built files being in the proper location"
$repo_path = "$PSScriptRoot\.."
New-Item -ItemType Directory -Path "$repo_path\Build" -Force
New-Item -ItemType Directory -Path "$repo_path\Build\dlls" -Force
Copy-Item "$repo_path\resource" "$repo_path\Build\resource" -Recurse
Copy-Item "$repo_path\install-update\install.ps1" "$repo_path\Build\install.ps1" -Force
Copy-Item "$repo_path\install-update\update.ps1" "$repo_path\Build\update.ps1" -Force
Copy-Item "$repo_path\PowerPointShell\bin\Release\*" "$repo_path\Build\dlls" -Recurse -Force
Remove-Item -path "$repo_path\build.zip" -Force
Compress-Archive -Path "$repo_path\Build\*" -DestinationPath "$repo_path\build.zip"
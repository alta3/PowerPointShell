# Ensure that we are running as an administrator
Start-Process powershell -Verb runAs

# Install required libs/mods for use from PowerShell directly
Install-PackageProvider Nuget –Force
Install-Module –Name PowerShellGet –Force
Install-Module -Name PowerShell-Yaml -Force

# Install to the users document folder under the name "Alta3 PowerPointShell"
New-Item -ItemType Directory -Path "$HOME\Documents\Alta3 PowerPointShell" -Force
Copy-Item $PSScriptRoot "$HOME\Documents\Alta3 PowerPointShell" -Recurse

# Change the environment Variable to look for modules located at this point as well
$CurrentEnvironment = [Environment]::GetEnvironmentVariable("PSModulePath", "Machine")
[Environment]::SetEnvironmentVariable("PSModulePath", $CurrentValue + ";$HOME\Documents\Alta3 PowerPointShell", "Machine")
Exit

# Ensure that we are running as an administrator
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

Write-Host "Installing Package Provider Nuget"
Install-PackageProvider Nuget -Force
Write-Host "Installing PowerShellGet Module"
Install-Module -Name PowerShellGet -Force
Write-Host "Installing PowerShell-Yaml Module"
Install-Module -Name PowerShell-Yaml -Force

# Install to the users document folder under the name "Alta3 PowerPointShell"
New-Item -ItemType Directory -Path "$HOME\Documents\Alta3 PowerPointShell" -Force
Copy-Item $PSScriptRoot\* "$HOME\Documents\Alta3 PowerPointShell" -Recurse
Write-Host "Installing files to $HOME\Documents\Alta3 PowerPointShell"

# Change the environment Variable to look for modules located at this point as well
$CurrentEnvironment = [Environment]::GetEnvironmentVariable("PSModulePath", "Machine")
[Environment]::SetEnvironmentVariable("PSModulePath", $CurrentValue + ";$HOME\Documents\Alta3 PowerPointShell", "Machine")
Write-Host "Added the environment variable to make the YAML modules persistently available!"
Write-Host "Installation successful!!!"
Read-Host "Press enter to continue"
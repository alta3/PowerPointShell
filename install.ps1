# Ensure that we are running as an administrator
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

# Install the neccessary packages to PowerShell To Install Modules and install the PowerShell-YAML Module
Write-Host "Installing Package Provider Nuget"
Install-PackageProvider Nuget -Force
Write-Host "Installing PowerShellGet Module"
Install-Module -Name PowerShellGet -Force
Write-Host "Installing PowerShell-Yaml Module"
Install-Module -Name PowerShell-Yaml -Force

# Install the PowerShell Module to a default module path accepted by PowerShell
Write-Host "Installing the PowerPointShell Module to the following location: $Env:ProgramFiles\WindowsPowerShell\Modules\PowerPointShell"
New-Item -ItemType Directory -Path $Env:ProgramFiles\WindowsPowerShell\Modules\PowerPointShell -Force
Copy-Item $PSScriptRoot\* "$Env:ProgramFiles\WindowsPowerShell\Modules\PowerPointShell" -Recurse -Force
Write-Host "Successfully added module to the module path!"

# Install to the users document folder under the name "Alta3 PowerPointShell"
Write-Host "Setting up working directory at the following location: $HOME\Documents\Alta3 PowerPointShell"
New-Item -ItemType Directory -Path "$HOME\Documents\Alta3 PowerPointShell" -Force
Copy-Item $PSScriptRoot\wrk "$HOME\Documents\Alta3 PowerPointShell" -Recurse -Force
Copy-Item $PSScriptRoot\pub "$HOME\Documents\Alta3 PowerPointShell" -Recurse -Force
Write-Host "Directory setup complete!"

# Let the user know how to access the commands
Write-Host "Installation successful!!!"
Write-Host "To see a list of available commands now directly accessible from powershell please type pps-help"
Write-Host "Typing `"pps-`" and attempting to tab complete should also work "
Read-Host "Press enter to continue"
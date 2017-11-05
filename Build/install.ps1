# Ensure that we are running as an administrator
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

# Install the neccessary packages to PowerShell To Install Modules and install the PowerShell-YAML Module
Install-PackageProvider Nuget -Force
Install-Module -Name PowerShellGet -Force
Install-Module -Name PowerShell-Yaml -Force

# Install the PowerShell Module to a default module path accepted by PowerShell
New-Item -ItemType Directory -Path $Env:ProgramFiles\WindowsPowerShell\Modules\PowerPointShell -Force
Copy-Item "$PSScriptRoot\dlls\*" "$Env:ProgramFiles\WindowsPowerShell\Modules\PowerPointShell" -Recurse -Force

# Install to the users document folder under the name "Alta3 PowerPoints"
New-Item -ItemType Directory -Path "$Env:UserProfile\Documents\Alta3 PowerPoints" -Force
New-Item -ItemType Directory -Path "$Env:UserProfile\Documents\Alta3 PowerPoints\sample" -Force  
New-Item -ItemType Directory -Path "$Env:UserProfile\Documents\Alta3 PowerPoints\working" -Force 
New-Item -ItemType Directory -Path "$Env:UserProfile\Documents\Alta3 PowerPoints\publish" -Force
New-Item -ItemType Directory -Path "$Env:UserProfile\Documents\Alta3 PowerPoints\resource" -Force 
Get-Item "$Env:UserProfile\Documents\Alta3 PowerPoints\resource" -Force | foreach { $_.Attributes = $_.Attributes -bor "Hidden" }
Copy-Item "$PSScriptRoot\resource\*" "$Env:UserProfile\Documents\Alta3 PowerPoints\resource" -Recurse -Force
Copy-Item "$PSScriptRoot\resource\mod.yml" "$Env:UserProfile\Documents\Alta3 PowerPoints\sample" -Force

# Let the user know how to access the commands
Write-Host "Installation complete!"
Write-Host "Newly generated PowerPoints will be saved to the following location: $Env:UserProfile\Documents\Alta3 PowerPoints"
Write-Host "To see a list of available commands now directly accessible from powershell please type A3-Help"
Write-Host "Tab completing after typing A3- will also show you the availble commands"
Write-Host "Get-Help will work on any of the commmands as well."
Write-Host "Press Enter to continue..."
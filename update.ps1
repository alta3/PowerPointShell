# Ensure we are running as an administrator!
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

# Update the module only!
Write-Host "Installing the PowerPointShell Module to the following location: $Env:ProgramFiles\WindowsPowerShell\Modules\PowerPointShell"
New-Item -ItemType Directory -Path $Env:ProgramFiles\WindowsPowerShell\Modules\PowerPointShell -Force
Copy-Item $PSScriptRoot\* "$Env:ProgramFiles\WindowsPowerShell\Modules\PowerPointShell" -Recurse -Force
Write-Host "Successfully added module to the module path!"
Read-Host "Press Enter To Continue"
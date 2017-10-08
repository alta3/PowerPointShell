# Main Entry Point For PowerPointShell
#
# Get the name of the subcommand
param(
    $cmd,
    [switch]$help
)
# Set default global variables
$cmds = New-Object System.Collections.ArrayList
$cmd_found = $false

# Loop through all available commands
foreach($_ in Get-ChildItem $PSScriptRoot\src\cmd -Name) {
   $cmd_name = [System.IO.Path]::GetFileNameWithoutExtension($_)
   $cmd_help = Get-Help "$PSScriptRoot\src\cmd\$cmd_name.ps1" -Full
   $cmds.Add($cmd_help.replace(".ps1","")) > $null
   if ($cmd -eq $cmd_name) {
       $cmd_found = $true
   }
}

if ($cmd_found -eq $true) {
    if ($help){
	write-host Retriving help!
        Get-Help "$PSScriptRoot\src\cmd\$cmd.ps1" -Full
    } else {
        & "$PSScriptRoot\src\cmd\$cmd.ps1" @args
    }
} else {
    Write-Host "Command not recognized!"
    Write-Host "The following is a list of available commands:"
    foreach ($_ in $cmds) {
        Write-Host -NoNewline $_
    }
}

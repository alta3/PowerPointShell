# Main Entry Point For PowerPointShell
#
# Get the name of the subcommand
param(
    $cmd,
    [switch]$help
)

# Test to see if supplied function exists
if (Test-Path "$PSScriptRoot\src\cmd\$cmd.ps1") { 
    if ($help) {
        Get-Help "$PSScriptRoot\src\cmd\$cmd.ps1"	    
    } else {
        & "$PSScriptRoot\src\cmd\$cmd.ps1" @args 
    }
} else {
    Write-Host "Command not recognized!"
    Write-Host "The following is a list of available commands:"
    foreach($_ in Get-ChildItem $PSScriptRoot\src\cmd -Name) {
        $cmd_help = Get-Help "$PSScriptRoot\src\cmd\$_" -Full
        $cmd_print = $cmd_help.replace(".ps1","")
	write-host -NoNewLine $cmd_print
    }
}

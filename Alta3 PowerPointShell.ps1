# Main Entry Point For PowerPointShell
function pps-help {
    Write-Host "Welcome to PowerPointShell!"
    Write-Host "The following is a list of the available commands provided by this module:"
    foreach($_ in Get-ChildItem $PSScriptRoot\src\cmd -Name) {
        $cmd_help = Get-Help "$PSScriptRoot\src\cmd\$_" -Full
        $cmd_print = $cmd_help.replace(".ps1","")
	    Write-Host -NoNewLine $cmd_print
    }
}

function pps-generate {
    [CmdletBinding()]     
    param(
        [string]$yamlpath,
        [switch]$help
    )
    if ($help) {
        Get-Help "$PSScriptRoort\src\cmd\gen.ps1" -Full
    } else {
        & "$PSScriptRoort\src\cmd\gen.ps1" -yamlpath $yamlpath
    }
}

function pps-list {}

function pps-add {
    param(
        [string]$message,
        [switch]$help
    )
    if ($help) {
        Get-Help "$PSScriptRoot\src\cmd\add.ps1" -Full
    } else {
        & "$PSScriptRoot\src\cmd\add.ps1" -message $message
    }
}

function pps-sync {}

function pps-update {}

pps-add -message "hello world"
pps
